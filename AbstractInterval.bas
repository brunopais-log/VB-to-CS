Attribute VB_Name = "AbstractInterval"
Option Explicit

Public Const isNoCalc As Integer = 0
Public Const isBeforeToAfter As Integer = 1
Public Const isBeforeToWithin As Integer = 2
Public Const isWithinToAfter As Integer = 3
Public Const isWithin As Integer = 4
Public Const isOutBefore As Integer = 5
Public Const isOutAfter As Integer = 6

Public Const iWorkedTimeMorning As Integer = 0
Public Const iWorkedTimeDay As Integer = 1
Public Const iWorkedTimeAfternoon As Integer = 2

Public dFullDispTimeAvailable As Date

Public MinimoTempoDISPEmRefeicao As Integer

Public TempoDescontoInicioCarreira As Integer

Public hasProcessError As Boolean
Public isAutoSave As Boolean
Public sError As String

Public Type DispPeriod
    dStartTime As Date
    dEndTime As Date
    iDispTimeMinutes As Double
    iDispTimeHours As Date
End Type

Public Type WorkTime
    dStartTime As Date
    dEndTime As Date
End Type

Public Type DriveTime
    dStartTime As Date
    dEndTime As Date
    isAU As Boolean
    isVazio As Boolean
    isInibeDisp As Boolean
End Type

Public Type DispByPeriod
    dStartTime As Date
    dEndTime As Date
    iDispTime As Long
    isEnabled As Boolean
    isInibeDisp As Boolean
End Type

Public Type AvaliatedWorkTimeDispTime
    iWorkedTimePeriod As Integer
    iWorkedTime As Double
    iDispTime As Double
    dWorkStartTime As Date
    dWorkEndTime As Date
    dDispStartTime As Date
    dDispEndTime As Date
End Type

Public Type CalculatedDispTime
    iDispTimeMinutes As Double
    iDispTimeHours As Date
    dDispStartTime As Date
    dDispEndTime As Date
    iTimePeriod As Integer
End Type

Public DP() As DispPeriod
Public WT() As WorkTime
Public dt() As DriveTime
Public DBP() As DispByPeriod
Public AWTDT() As AvaliatedWorkTimeDispTime
Public CDT() As CalculatedDispTime

Public Inc_DP As Integer
Public Inc_WT As Integer
Public Inc_DT As Integer
Public Inc_DBP As Integer
Public Inc_AWTDT As Integer
Public Inc_CDT As Integer

Const Max4h30m As Integer = 270
'Const MaxTimeBetweenDT As Integer = 20
Public MaxTimeBetweenDT As Integer

Dim TimeToNextPeriod As Long
Dim iLastiIntervalType As Long

'**************************************************
'DRIVING PAUSES
'**************************************************
Public Type Pauses
    isEnablePause As Boolean
    isCarreiraPlus49Kms As Boolean
    isAluguerPlus49Kms As Boolean
    isOcasional As Boolean
    iTotalDrivingTime As Double
End Type

Public Type Pause
    dStartTime As Date
    dEndTime As Date
    iTotalPauseTime As Double
End Type

Public DrivingPauses() As Pauses
Public DrivingPause() As Pause
Public Inc_Pauses As Integer
Public Inc_Pause As Integer

Public isForce2Pauses As Boolean
Public isForceSinglePause As Boolean
'**************************************************


'****************************************************************************************************************
'INITIALIZE
'****************************************************************************************************************
Public Sub InitializeProcess()
On Error GoTo ErrHandler

    ReDim DP(0)
    ReDim WT(0)
    ReDim dt(0)
    ReDim DBP(0)
    ReDim AWTDT(0)
    ReDim CDT(0)
    ReDim DrivingPauses(0)
    ReDim DrivingPause(0)
    
    Inc_DP = 0
    Inc_WT = 0
    Inc_DT = 0
    Inc_DBP = 0
    Inc_AWTDT = 0
    Inc_CDT = 0
    Inc_Pauses = 0
    Inc_Pause = 0

Exit Sub
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:InitializeProcess", Err.Description
End Sub
'****************************************************************************************************************

'****************************************************************************************************************
'ADD INTERVAL
'****************************************************************************************************************
Public Function addTimeInterval_DispPeriod(ByVal dStart_Time As Date, ByVal dEnd_Time As Date, ByVal isClear As Boolean) As Boolean
On Error GoTo ErrHandler

    If isClear = True Then
        Inc_DP = 0
        ReDim DP(Inc_DP)
    End If

    If DateDiff("n", dStart_Time, dEnd_Time) < 0 Then
        If isAutoSave = False Then
            MsgBox "A data de inicio n�o pode ser superior � data de fim!", vbCritical + vbOKOnly
            sError = "addTimeInterval_DispPeriod: A data de inicio n�o pode ser superior � data de fim!"
        End If
        hasProcessError = True
        addTimeInterval_DispPeriod = False
    Else
        Inc_DP = Inc_DP + 1
        ReDim Preserve DP(Inc_DP)
        
        With DP(Inc_DP - 1)
            .dStartTime = dStart_Time
            .dEndTime = dEnd_Time
        End With
        
        addTimeInterval_DispPeriod = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:addTimeInterval_DispPeriod", Err.Description
    hasProcessError = True
    sError = "addTimeInterval_DispPeriod: " & Err.Description
End Function

Public Function addTimeInterval_WorkTime(ByVal dStart_Time As Date, ByVal dEnd_Time As Date, ByVal isClear As Boolean) As Boolean
On Error GoTo ErrHandler

    If isClear = True Then
        Inc_WT = 0
        ReDim WT(Inc_DP)
    End If
    
    If DateDiff("n", dStart_Time, dEnd_Time) < 0 Then
        If isAutoSave = False Then
            MsgBox "A data de inicio n�o pode ser superior � data de fim!", vbCritical + vbOKOnly
            sError = "addTimeInterval_WorkTime: A data de inicio n�o pode ser superior � data de fim!"
        End If
        hasProcessError = True
        addTimeInterval_WorkTime = False
    Else
        Inc_WT = Inc_WT + 1
        ReDim Preserve WT(Inc_WT)
        
        With WT(Inc_WT - 1)
            .dStartTime = dStart_Time
            .dEndTime = dEnd_Time
        End With
        
        addTimeInterval_WorkTime = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:addTimeInterval_WorkTime", Err.Description
    hasProcessError = True
    sError = "addTimeInterval_WorkTime: " & Err.Description
End Function

Public Function addTimeInterval_DriveTime(ByVal dStart_Time As Date, ByVal dEnd_Time As Date, ByVal isClear As Boolean, ByVal isFirst As Boolean, ByVal isLast As Boolean, Optional ByVal isAU As Boolean = False, Optional ByVal isVazio As Boolean = False, Optional ByVal isInibirDisponibilidade As Boolean = False) As Boolean
    Dim iAnteciparEntrada As Double
On Error GoTo ErrHandler

    If isClear = True Then
        Inc_DT = 0
        ReDim dt(Inc_DT)
    End If
    
    If DateDiff("n", dStart_Time, dEnd_Time) < 0 Then
        If isAutoSave = False Then
            MsgBox "A data de inicio n�o pode ser superior � data de fim!", vbCritical + vbOKOnly
            sError = "addTimeInterval_DriveTime: A data de inicio n�o pode ser superior � data de fim!"
        End If
        hasProcessError = True
        addTimeInterval_DriveTime = False
    Else
        If TempoDescontoInicioCarreira > 0 Then
            'REMOVED - Requested by David Alaba�a
            'And isVazio = False Then
            'REMOVES TIME TO THE START TIME
            
            If oFilial.id_Filial = 1 Then
                If isFirst = True Then
                    '*********************************************************************************
                    '   COLOCADA ESTA OP��O DEVIDO AO FACTO DA ESTREMADURA FUNCIONAR DE MODO DIFERENTE
                    '*********************************************************************************
                    iAnteciparEntrada = TempoDescontoInicioCarreira * (-1)
                    dStart_Time = DateAdd("n", iAnteciparEntrada, dStart_Time)
                    If isLast = True Then
                        dEnd_Time = DateAdd("n", TempoDescontoInicioCarreira, dEnd_Time)
                    Else
                        dEnd_Time = DateAdd("n", oParametrosGerais.MinutosIntervalo, dEnd_Time)
                    End If
                    '*********************************************************************************
                    '*********************************************************************************
                Else
                    '*********************************************************************************
                    '   ORIGINAL
                    '*********************************************************************************
                    iAnteciparEntrada = oParametrosGerais.MinutosIntervalo * (-1)
                    dStart_Time = DateAdd("n", iAnteciparEntrada, dStart_Time)
                    If isLast = True Then
                        dEnd_Time = DateAdd("n", TempoDescontoInicioCarreira, dEnd_Time)
                    Else
                        dEnd_Time = DateAdd("n", iAnteciparEntrada, dEnd_Time)
                    End If
                    '*********************************************************************************
                    '*********************************************************************************
                End If
            Else
                iAnteciparEntrada = TempoDescontoInicioCarreira * (-1)
                dStart_Time = DateAdd("n", iAnteciparEntrada, dStart_Time)
                dEnd_Time = DateAdd("n", TempoDescontoInicioCarreira, dEnd_Time)
            End If
        End If
    
        Inc_DT = Inc_DT + 1
        ReDim Preserve dt(Inc_DT)
        
        With dt(Inc_DT - 1)
            .dStartTime = dStart_Time
            .dEndTime = dEnd_Time
            .isAU = isAU
            .isInibeDisp = isInibirDisponibilidade
        End With
        
        addTimeInterval_DriveTime = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:addTimeInterval_DriveTime", Err.Description
    hasProcessError = True
    sError = "addTimeInterval_DriveTime: " & Err.Description
End Function

Private Function addTimeInterval_DispByPeriod(ByVal dStart_Time As Date, ByVal dEnd_Time As Date, ByVal isClear As Boolean, ByVal isInibeDisp As Boolean) As Boolean
On Error GoTo ErrHandler

    If isClear = True Then
        Inc_DBP = 0
        ReDim DBP(Inc_DBP)
    End If
    
    If DateDiff("n", dStart_Time, dEnd_Time) < 0 Then
        If isAutoSave = False Then
            MsgBox "A data de inicio n�o pode ser superior � data de fim!", vbCritical + vbOKOnly
            sError = "addTimeInterval_DispByPeriod: A data de inicio n�o pode ser superior � data de fim!"
        End If
        hasProcessError = True
        addTimeInterval_DispByPeriod = False
    Else
        Inc_DBP = Inc_DBP + 1
        ReDim Preserve DBP(Inc_DBP)
        
        With DBP(Inc_DBP - 1)
            .dStartTime = dStart_Time
            .dEndTime = dEnd_Time
            .isInibeDisp = isInibeDisp
        End With
        
        addTimeInterval_DispByPeriod = True
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:addTimeInterval_DispByPeriod", Err.Description
    hasProcessError = True
    sError = "addTimeInterval_DispByPeriod: " & Err.Description
End Function
'****************************************************************************************************************

'****************************************************************************************************************
'CONTAINS FUNCTIONS
'****************************************************************************************************************
'Private Function ContainsValueTime(ByVal dTime As Date) As Integer
'    Dim iResult As Integer
'On Error GoTo ErrHandler
'
'    If DateDiff("n", dTime, dStartTime) <= 0 And DateDiff("n", dTime, dEndTime) >= 0 Then
'        iResult = isWithin
'    ElseIf DateDiff("n", dTime, dEndTime) < 0 Then
'        iResult = isOutAfter
'    ElseIf DateDiff("n", dTime, dStartTime) > 0 Then
'        iResult = isOutBefore
'    Else
'        iResult = isNoCalc
'    End If
'
'    ContainsValueTime = iResult
'
'Exit Function
'ErrHandler:
'    ContainsValueTime = isNoCalc
'    Err.Raise Err.Number, "AbstractInterval:ContainsValueTime", Err.Description
'End Function

Public Function ContainsIntervalTime(ByVal dStartWT As Date, ByVal dEndWT As Date, ByVal dStartDP As Date, ByVal dEndDP As Date) As Integer
    Dim iResult As Integer
On Error GoTo ErrHandler

    If (DateDiff("n", dStartWT, dStartDP) > 0) And (DateDiff("n", dEndWT, dEndDP) < 0) Then
        iResult = isBeforeToAfter
    ElseIf (DateDiff("n", dStartWT, dStartDP) > 0) And (DateDiff("n", dEndWT, dStartDP) <= 0 And DateDiff("n", dEndWT, dEndDP) >= 0) Then
        iResult = isBeforeToWithin
    ElseIf (DateDiff("n", dStartWT, dStartDP) <= 0) And (DateDiff("n", dStartWT, dEndDP) >= 0 And DateDiff("n", dEndWT, dEndDP) < 0) Then
        iResult = isWithinToAfter
    ElseIf (DateDiff("n", dStartWT, dStartDP) <= 0) And (DateDiff("n", dEndWT, dEndDP) >= 0) Then
            iResult = isWithin
    ElseIf (DateDiff("n", dStartWT, dStartDP) > 0) And (DateDiff("n", dEndWT, dStartDP) > 0 And DateDiff("n", dEndWT, dEndDP) > 0) Then
        iResult = isOutBefore
    ElseIf (DateDiff("n", dStartWT, dEndDP) < 0) And (DateDiff("n", dEndWT, dEndDP) < 0) Then
        iResult = isOutAfter
    Else
        iResult = isNoCalc
    End If

    ContainsIntervalTime = iResult

Exit Function
ErrHandler:
    ContainsIntervalTime = isNoCalc
    Err.Raise Err.Number, "AbstractInterval:ContainsIntervalTime", Err.Description
    hasProcessError = True
    sError = "ContainsIntervalTime: " & Err.Description
End Function
'****************************************************************************************************************

'****************************************************************************************************************
'GET INFO FUNCTIONS
'****************************************************************************************************************
Public Function GetCalculatedDispTime() As Date
    Dim Inc As Integer
    Dim dTotalTime As Double
On Error GoTo ErrHandler

    For Inc = 0 To Inc_DBP - 1
        If DBP(Inc).isEnabled = True Then
            dTotalTime = dTotalTime + DBP(Inc).iDispTime
        End If
    Next Inc
    
    Dim sDispTime As String
    Dim sSplit() As String
    sSplit() = Split(CStr(dTotalTime), ",")
    If UBound(sSplit()) > 0 Then
        If IsDate(ElapsedTimeDouble(sSplit())) = True Then
            sDispTime = CDate(ElapsedTimeDouble(sSplit()))
        Else
            sDispTime = CDate(0)
        End If
    Else
        If CDbl(sSplit(0)) <= 86400 Then
            sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
        Else
            sDispTime = CDate(0)
        End If
    End If
    
    GetCalculatedDispTime = CDate(sDispTime)

Exit Function
ErrHandler:
    If Err.Number = 6 Then ' overflow
    Else
        Err.Raise Err.Number, "AbstractInterval:GetCalculatedDispTime", Err.Description
        GetCalculatedDispTime = CDate("00:00")
        hasProcessError = True
        sError = "GetCalculatedDispTime: " & Err.Description
    End If
End Function

Public Function GetCalculatedDispTimeInMinutes() As Double
    Dim Inc As Integer
    Dim dTotalTime As Double
On Error GoTo ErrHandler

    For Inc = 0 To Inc_DBP - 1
        If DBP(Inc).isEnabled = True Then
            dTotalTime = dTotalTime + DBP(Inc).iDispTime
        End If
    Next Inc
    
    GetCalculatedDispTimeInMinutes = dTotalTime

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetCalculatedDispTimeInMinutes", Err.Description
    GetCalculatedDispTimeInMinutes = 0
    hasProcessError = True
    sError = "GetCalculatedDispTimeInMinutes: " & Err.Description
End Function

Public Function GetFullWorkedTime() As Date
    Dim sDispTime As String
    Dim sSplit() As String
    Dim Inc As Integer
    Dim dTotalTime As Double
On Error GoTo ErrHandler

    For Inc = 0 To Inc_WT - 1
        dTotalTime = dTotalTime + DateDiff("n", WT(Inc).dStartTime, WT(Inc).dEndTime)
    Next Inc
    
    sSplit() = Split(CStr(dTotalTime), ",")
    If UBound(sSplit()) > 0 Then
        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
    Else
        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
    End If

    GetFullWorkedTime = CDate(sDispTime)

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetFullWorkedTime", Err.Description
    GetFullWorkedTime = CDate("00:00")
    hasProcessError = True
    sError = "GetFullWorkedTime: " & Err.Description
End Function

Public Function GetWorkedTime(ByVal dNormalWorkTime As Date) As Date
    Dim sDispTime As String
    Dim sSplit() As String
    Dim Inc As Integer
    Dim dTotalTime As Double
On Error GoTo ErrHandler

    For Inc = 0 To Inc_WT - 1
        dTotalTime = dTotalTime + DateDiff("n", WT(Inc).dStartTime, WT(Inc).dEndTime)
    Next Inc
    
    dTotalTime = dTotalTime - (dNormalWorkTime * 60 * 24)
    If dTotalTime < 0 Then
        dTotalTime = 0
    End If
    
    sSplit() = Split(CStr(dTotalTime), ",")
    If UBound(sSplit()) > 0 Then
        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
    Else
        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
    End If

    GetWorkedTime = CDate(sDispTime)

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetWorkedTime", Err.Description
    GetWorkedTime = CDate("00:00")
    hasProcessError = True
    sError = "GetWorkedTime: " & Err.Description
End Function

Public Function GetCalculatedWorkedTime(ByVal dNormalWorkTime As Date) As Date
    Dim sDispTime As String
    Dim sSplit() As String
    Dim Inc As Integer
    Dim dTotalTime As Double
    Dim dNormalWorkTime_Support As Double
On Error GoTo ErrHandler
    
    For Inc = 0 To Inc_WT - 1
        If CStr(WT(Inc).dStartTime) <> "00:00:00" And CStr(WT(Inc).dEndTime) <> "00:00:00" Then
            dTotalTime = dTotalTime + DateDiff("n", WT(Inc).dStartTime, WT(Inc).dEndTime)
        End If
    Next Inc
    
    dNormalWorkTime_Support = (dNormalWorkTime * 60 * 24)
    dTotalTime = dTotalTime - dNormalWorkTime_Support
    If dTotalTime < 0 Then
        dTotalTime = 0
    End If
    
    sSplit() = Split(CStr(dTotalTime), ",")
    If UBound(sSplit()) > 0 Then
        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
    Else
        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
    End If

    FinalyzeDispTimeCalculation (CDate(sDispTime))
    CalculateDispTimeByPeriod

    GetCalculatedWorkedTime = CDate(sDispTime)

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetCalculatedWorkedTime", Err.Description
    GetCalculatedWorkedTime = CDate("00:00")
    hasProcessError = True
    sError = "GetCalculatedWorkedTime: " & Err.Description
End Function
'****************************************************************************************************************

'****************************************************************************************************************
' WORKING ON DETAILS
'****************************************************************************************************************
Public Function GetDriveTime() As Long
    Dim Inc As Integer
    Dim isFirst As Boolean
    Dim lngDriveTime As Long
On Error GoTo ErrHandler
    isFirst = True
    
    For Inc = 0 To Inc_DT - 1
        If Inc < Inc_DT - 1 Then
            lngDriveTime = lngDriveTime + CDbl(DateDiff("n", dt(Inc).dStartTime, dt(Inc).dEndTime))
        End If
    Next Inc
    
    GetDriveTime = lngDriveTime

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetDispTime", Err.Description
    hasProcessError = True
    sError = "GetDriveTime: " & Err.Description
End Function

Public Function GetDispTime() As Boolean
    Dim Inc As Integer
    Dim isFirst As Boolean
    Dim isFirstOf2 As Boolean
    Dim isDonePause As Boolean
    Dim dtNewStartDate As Date
    Dim isPauseFinished As Boolean
    Dim isPauseBlock As Boolean
    Dim dDataINI As Date
    Dim dDataEnd As Date
    Dim dDataIniNext As Date
    Dim Inc2 As Integer
    Dim isCalcOK As Boolean
    Dim isBypassError As Boolean
On Error GoTo ErrHandler
    isFirst = True
    isFirstOf2 = True
    
    If isPauseBlock = False And DrivingPauses(0).isEnablePause = True Then
        If isForceSinglePause = True Then
            If DateDiff("n", WT(0).dStartTime, WT(0).dEndTime) <= Max4h30m Then
                For Inc = 0 To Inc_WT - 1
                    If DateDiff("n", WT(Inc).dEndTime, WT(Inc + 1).dStartTime) >= 45 And isDonePause = False Then
                        If AddPause(WT(Inc).dEndTime, WT(Inc + 1).dStartTime) = False Then
                            Exit Function
                        Else
                            isDonePause = True
                            isPauseFinished = True
                            isPauseBlock = True
                            Exit For
                        End If
                    End If
                Next Inc
            End If
        End If
    End If
    
    '**********************************************************************************************
    'INSERTED AT REQUEST FROM David Alaba�a
    '**********************************************************************************************
    If DateDiff("n", WT(0).dStartTime, dt(0).dStartTime) > 0 Then
        If addTimeInterval_DispByPeriod(WT(0).dStartTime, dt(0).dStartTime, isFirst, dt(0).isInibeDisp) = False Then
            Exit Function
        Else
            DBP(Inc_DBP - 1).iDispTime = DateDiff("n", DBP(Inc_DBP - 1).dStartTime, DBP(Inc_DBP - 1).dEndTime)
            If DBP(Inc_DBP - 1).iDispTime >= MaxTimeBetweenDT Then
                DBP(Inc_DBP - 1).isEnabled = True
            Else
                DBP(Inc_DBP - 1).isEnabled = False
            End If
            
            isFirst = False
        End If
    Else
        If addTimeInterval_DispByPeriod(WT(0).dStartTime, WT(0).dStartTime, isFirst, dt(0).isInibeDisp) = False Then
            Exit Function
        Else
            DBP(Inc_DBP - 1).iDispTime = DateDiff("n", DBP(Inc_DBP - 1).dStartTime, DBP(Inc_DBP - 1).dEndTime)
            If DBP(Inc_DBP - 1).iDispTime >= MaxTimeBetweenDT Then
                DBP(Inc_DBP - 1).isEnabled = True
            Else
                DBP(Inc_DBP - 1).isEnabled = False
            End If
            
            isFirst = False
        End If
    End If
    '**********************************************************************************************
    
    If DBP(Inc_DBP - 1).isEnabled = True And DBP(Inc_DBP - 1).isInibeDisp = True Then
        DBP(Inc_DBP - 1).isEnabled = False
    End If

    For Inc = 0 To Inc_DT - 1
        If Inc < Inc_DT - 1 Then
            If DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime) > 0 Then
                If addTimeInterval_DispByPeriod(dt(Inc).dEndTime, dt(Inc + 1).dStartTime, isFirst, dt(Inc + 1).isInibeDisp) = False Then
                    Exit Function
                End If
                '**********************************************************************************************
                'NORMAL PROCESS - OLD PROCESS
                '**********************************************************************************************
'                With DBP(Inc_DBP - 1)
'                    .iDispTime = DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime)
'                    If .iDispTime >= MaxTimeBetweenDT Then
'                        .isEnabled = True
'                    Else
'                        .isEnabled = False
'                    End If
'                End With
                '**********************************************************************************************
                'NORMAL PROCESS - NEW PROCESS
                '**********************************************************************************************
                If isPauseBlock = False And DrivingPauses(0).isEnablePause = True Then
                    If isForce2Pauses = True Then
                        If DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime) >= 15 And isFirstOf2 = True And isDonePause = False Then
                            '**********************************************************************************
                            ' VALIDATE IF IT IS AFTER LUNCH
                            '**********************************************************************************
LunchAfter4H30M:
                            isCalcOK = False
                            
                            For Inc2 = 0 To Inc_WT - 1
                                If DateDiff("n", WT(Inc2).dStartTime, WT(Inc2).dEndTime) > 0 And DateDiff("n", WT(Inc2).dStartTime, WT(Inc2).dEndTime) <= Max4h30m Then
                                    If DateDiff("n", dt(Inc).dEndTime, WT(Inc2).dEndTime) < 0 Then
                                        dDataINI = WT(Inc2).dStartTime
                                        dDataEnd = WT(Inc2).dEndTime
                                        If Inc2 < 4 Then
                                            If DateDiff("n", WT(Inc2 + 1).dStartTime, WT(Inc2 + 1).dEndTime) > 0 Then
                                                dDataIniNext = WT(Inc2 + 1).dStartTime
                                            End If
                                        Else
                                            isCalcOK = False
                                            Exit For
                                        End If
                                        isCalcOK = True
                                    End If
                                Else
                                    Exit For
                                End If
                            Next Inc2
                            
                            If isCalcOK = True And DateDiff("n", WT(Inc2).dStartTime, WT(Inc2).dEndTime) <= Max4h30m Then
                                isFirstOf2 = False
                                If AddPause(dDataEnd, dDataIniNext) = False Then
                                    Exit Function
                                End If
                            ElseIf isCalcOK = False And DateDiff("n", WT(Inc2).dStartTime, WT(Inc2).dEndTime) <= Max4h30m Then
                                isFirstOf2 = False
                                If AddPause(dt(Inc).dEndTime, dt(Inc + 1).dStartTime) = False Then
                                    Exit Function
                                End If
                            End If
                            
                            '**********************************************************************************
'                           DISABLED FOR RESOLVING 15MIN ON END ISSUE
'                            isFirstOf2 = False
'                            If AddPause(dt(Inc).dEndTime, dt(Inc + 1).dStartTime) = False Then
'                                Exit Function
'                            End If
                            If DateDiff("n", WT(Inc2).dStartTime, WT(Inc2).dEndTime) <= Max4h30m Then
                                dDataINI = DrivingPause(Inc_Pause - 1).dStartTime
                                dDataEnd = DrivingPause(Inc_Pause - 1).dEndTime
'                            If MatchPAUSETimeWithWorkTime(dDataIni, dDataEnd) = False Then
'                                If DateDiff("n", dDataIni, dDataEnd) > 15 Then
'                                    dDataEnd = DateAdd("n", 15, dDataIni)
'                                ElseIf DateDiff("n", dDataIni, dDataEnd) < 15 Then
'                                    'REMOVE RECORD
'                                    Inc_Pause = Inc_Pause - 1
'                                    ReDim Preserve DrivingPause(Inc_Pause)
'                                    DrivingPause(Inc_Pause - 1).dStartTime = "00:00"
'                                    DrivingPause(Inc_Pause - 1).dEndTime = "00:00"
'                                    DrivingPause(Inc_Pause - 1).iTotalPauseTime = 0
'                                ElseIf DateDiff("n", dDataIni, dDataEnd) = 15 Then
'                                    DrivingPause(Inc_Pause - 1).dStartTime = dDataIni
'                                    DrivingPause(Inc_Pause - 1).dEndTime = dDataEnd
'
'                                    dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
'                                    isDonePause = True
'                                End If
'                            Else
'                                dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
'                                isDonePause = True
'                            End If
                            
                                dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
                                isDonePause = True
                            End If
                        ElseIf DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime) >= MaxTimeBetweenDT And isFirstOf2 = False And isDonePause = False Then
                            For Inc2 = 0 To Inc_WT - 1
                                If DateDiff("n", WT(Inc2).dEndTime, WT(Inc2 + 1).dStartTime) >= MaxTimeBetweenDT And isDonePause = False Then
                                    If DateDiff("n", WT(Inc2).dStartTime, WT(Inc2).dEndTime) >= Max4h30m Then
                                        GoTo LunchAfter4H30M
                                    End If
                                    If AddPause(WT(Inc2).dEndTime, WT(Inc2 + 1).dStartTime) = False Then
                                        isBypassError = True
                                        Exit For
                                    Else
                                        isDonePause = True
                                        isPauseFinished = True
                                        Exit For
                                    End If
                                End If
                            Next Inc2
                        
                            If isBypassError = False Then
                                If isDonePause = False And isPauseFinished = False Then
                                    If AddPause(dt(Inc).dEndTime, dt(Inc + 1).dStartTime) = False Then
                                        Exit Function
                                    End If
                                End If
                                
                                dDataINI = DrivingPause(Inc_Pause - 1).dStartTime
                                dDataEnd = DrivingPause(Inc_Pause - 1).dEndTime
    '                            If MatchPAUSETimeWithWorkTime(dDataIni, dDataEnd) = False Then
    '                                If DateDiff("n", dDataIni, dDataEnd) > MaxTimeBetweenDT Then
    '                                    dDataEnd = DateAdd("n", MaxTimeBetweenDT, dDataIni)
    '                                ElseIf DateDiff("n", dDataIni, dDataEnd) < MaxTimeBetweenDT Then
    '                                    'REMOVE RECORD
    '                                    Inc_Pause = Inc_Pause - 1
    '                                    ReDim Preserve DrivingPause(Inc_Pause)
    '                                    DrivingPause(Inc_Pause).dStartTime = "00:00"
    '                                    DrivingPause(Inc_Pause).dEndTime = "00:00"
    '                                    DrivingPause(Inc_Pause).iTotalPauseTime = 0
    '                                ElseIf DateDiff("n", dDataIni, dDataEnd) = MaxTimeBetweenDT Then
    '                                    DrivingPause(Inc_Pause - 1).dStartTime = dDataIni
    '                                    DrivingPause(Inc_Pause - 1).dEndTime = dDataEnd
    '                                    DrivingPause(Inc_Pause - 1).iTotalPauseTime = DateDiff("n", DrivingPause(Inc_Pause - 1).dStartTime, DrivingPause(Inc_Pause - 1).dEndTime)
    '
    '                                    dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
    '                                    isDonePause = True
    '                                    isPauseFinished = True
    '                                End If
    '                            Else
                                    dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
                                    isDonePause = True
                                    isPauseFinished = True
    '                            End If
'                            End If
                        End If
                    ElseIf isForceSinglePause = True Then
                        If DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime) >= 45 And isDonePause = False Then
                            If DateDiff("n", WT(0).dStartTime, dt(Inc).dEndTime) <= Max4h30m Then
                                If AddPause(dt(Inc).dEndTime, dt(Inc + 1).dStartTime) = False Then
                                    Exit Function
                                End If
                                
                                dDataINI = DrivingPause(Inc_Pause - 1).dStartTime
                                dDataEnd = DrivingPause(Inc_Pause - 1).dEndTime
    '                            If MatchPAUSETimeWithWorkTime(dDataIni, dDataEnd) = False Then
    '                                If DateDiff("n", dDataIni, dDataEnd) > 45 Then
    '                                    dDataEnd = DateAdd("n", 45, dDataIni)
    '                                ElseIf DateDiff("n", dDataIni, dDataEnd) < 45 Then
    '                                    'REMOVE RECORD
    '                                    Inc_Pause = Inc_Pause - 1
    '                                    ReDim Preserve DrivingPause(Inc_Pause)
    '                                    If Inc_Pause > 0 Then
    '                                        DrivingPause(Inc_Pause - 1).dStartTime = "00:00"
    '                                        DrivingPause(Inc_Pause - 1).dEndTime = "00:00"
    '                                        DrivingPause(Inc_Pause - 1).iTotalPauseTime = 0
    '                                    Else
    '                                        DrivingPause(Inc_Pause).dStartTime = "00:00"
    '                                        DrivingPause(Inc_Pause).dEndTime = "00:00"
    '                                        DrivingPause(Inc_Pause).iTotalPauseTime = 0
    '                                    End If
    '                                ElseIf DateDiff("n", dDataIni, dDataEnd) = 45 Then
    '                                    DrivingPause(Inc_Pause - 1).dStartTime = dDataIni
    '                                    DrivingPause(Inc_Pause - 1).dEndTime = dDataEnd
    '
    '                                    dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
    '                                    isDonePause = True
    '                                    isPauseFinished = True
    '                                End If
    '                            Else
                                    dtNewStartDate = DrivingPause(Inc_Pause - 1).dStartTime
                                    isDonePause = True
                                    isPauseFinished = True
                                End If
                            End If
                        End If
                    End If
                End If
                
                '**********************************************************************************
                'TRIES TO GIVE PAUSE IF LUNCH TIME AND NO PAUSE GIVEN ISSUE
                '**********************************************************************************
                If isForce2Pauses = True And isFirstOf2 = True Then
                    For Inc2 = 0 To Inc_WT - 1
                        If isFirstOf2 = True And DateDiff("n", WT(Inc2).dEndTime, dt(Inc).dEndTime) <= 0 Then
                            isFirstOf2 = False
                            If AddPause(WT(Inc2).dEndTime, WT(Inc2 + 1).dStartTime) = False Then
                                Exit Function
                            End If
                        End If
                    Next Inc2
                End If
                '**********************************************************************************
                            
                
                With DBP(Inc_DBP - 1)
                    If isDonePause = True And isPauseBlock = False Then
                        If DateDiff("n", dt(Inc).dEndTime, DrivingPause(Inc_Pause - 1).dEndTime) >= 0 Then
                            .iDispTime = DateDiff("n", DrivingPause(Inc_Pause - 1).dEndTime, dt(Inc + 1).dStartTime)
                            '.iDispTime = DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime)
                            .dStartTime = DrivingPause(Inc_Pause - 1).dEndTime
                            
                            If isPauseFinished = True Then
                                isPauseBlock = True
                            End If
                            isDonePause = False
                        Else
                            '********************************************************************
                            ' BUG THAT SET'S DISP ON DRIVE TIME WHEN PAUSE ON 15/30
                            '********************************************************************
                            If isPauseFinished = True Then
                                isPauseBlock = True
                            End If
                            isDonePause = False
                            
                            .iDispTime = DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime)
                            '********************************************************************
                        End If
                    Else
                        .iDispTime = DateDiff("n", dt(Inc).dEndTime, dt(Inc + 1).dStartTime)
                    End If
                        
                    If .iDispTime >= MaxTimeBetweenDT Then
                        .isEnabled = True
                        
                        If .isEnabled = True And .isInibeDisp = True Then
                            .isEnabled = False
                        End If
                    Else
                        .isEnabled = False
                    End If
                End With
                '**********************************************************************************************
                isFirst = False
            End If
        End If
    Next Inc
    
    '**********************************************************************************
    'TRIES TO GIVE PAUSE IF LUNCH TIME AND NO PAUSE GIVEN ISSUE
    '**********************************************************************************
'    For Inc2 = 0 To Inc_WT - 1
'        If isFirstOf2 = True And DateDiff("n", WT(Inc2).dEndTime, dt(Inc).dEndTime) <= 0 Then
'            isFirstOf2 = False
'            If AddPause(WT(Inc2).dEndTime, WT(Inc2 + 1).dStartTime) = False Then
'                Exit Function
'            End If
'        End If
'    Next Inc2
    '**********************************************************************************

    '**********************************************************************************************
    'INSERTED AT REQUEST FROM David Alaba�a
    '**********************************************************************************************
    If WT(Inc_WT - 1).dStartTime <> WT(Inc_WT - 1).dEndTime Then
        dDataEnd = WT(Inc_WT - 1).dEndTime
    ElseIf WT(Inc_WT - 2).dStartTime <> WT(Inc_WT - 2).dEndTime Then
        dDataEnd = WT(Inc_WT - 2).dEndTime
    ElseIf WT(Inc_WT - 3).dStartTime <> WT(Inc_WT - 3).dEndTime Then
        dDataEnd = WT(Inc_WT - 3).dEndTime
    ElseIf WT(Inc_WT - 4).dStartTime <> WT(Inc_WT - 4).dEndTime Then
        dDataEnd = WT(Inc_WT - 4).dEndTime
    Else
        GoTo Bypass
    End If
    
    If DateDiff("n", dt(Inc_DT - 1).dEndTime, dDataEnd) > 0 Then
        If addTimeInterval_DispByPeriod(dt(Inc_DT - 1).dEndTime, dDataEnd, isFirst, dt(Inc_DT - 1).isInibeDisp) = False Then
            Exit Function
        Else
            DBP(Inc_DBP - 1).iDispTime = DateDiff("n", DBP(Inc_DBP - 1).dStartTime, DBP(Inc_DBP - 1).dEndTime)
            If DBP(Inc_DBP - 1).iDispTime >= MaxTimeBetweenDT Then
                DBP(Inc_DBP - 1).isEnabled = True
                
                If DBP(Inc_DBP - 1).isEnabled = True And DBP(Inc_DBP - 1).isInibeDisp = True Then
                    DBP(Inc_DBP - 1).isEnabled = False
                End If
            Else
                DBP(Inc_DBP - 1).isEnabled = False
            End If
            
            isFirst = False
        End If
    Else
        If addTimeInterval_DispByPeriod(dDataEnd, dDataEnd, isFirst, False) = False Then
            Exit Function
        Else
            DBP(Inc_DBP - 1).iDispTime = DateDiff("n", DBP(Inc_DBP - 1).dStartTime, DBP(Inc_DBP - 1).dEndTime)
            If DBP(Inc_DBP - 1).iDispTime >= MaxTimeBetweenDT Then
                DBP(Inc_DBP - 1).isEnabled = True
                
                If DBP(Inc_DBP - 1).isEnabled = True And DBP(Inc_DBP - 1).isInibeDisp = True Then
                    DBP(Inc_DBP - 1).isEnabled = False
                End If
            Else
                DBP(Inc_DBP - 1).isEnabled = False
            End If
            
            isFirst = False
        End If
    End If
Bypass:
    '**********************************************************************************************

    
    '******************************************************************************************************************************
    '                           CREATED FOR RESOLVING 30MIN ON END ISSUE
    '******************************************************************************************************************************
    If isPauseBlock = False And (isForce2Pauses = True Or isForceSinglePause = True) And DrivingPauses(0).isEnablePause = True Then
        If isForceSinglePause = False Then
            isCalcOK = False
            For Inc = 0 To Inc_WT - 1
                For Inc2 = 0 To Inc_Pause - 1
                    If WT(Inc).dEndTime = DrivingPause(Inc2).dStartTime Then
                        isCalcOK = False
                    Else
                        isCalcOK = True
                    End If
                Next Inc2
                
                If isCalcOK = True Then
                    If DateDiff("n", WT(0).dStartTime, WT(Inc).dEndTime) <= Max4h30m Then
                        If DateDiff("n", WT(Inc).dEndTime, WT(Inc + 1).dStartTime) >= MaxTimeBetweenDT And isDonePause = False Then
                            If AddPause(WT(Inc).dEndTime, WT(Inc + 1).dStartTime) = False Then
                                Exit Function
                            Else
                                isDonePause = True
                                isPauseFinished = True
                                isPauseBlock = True
                                Exit For
                            End If
                        End If
                    End If
                Else
                    If DateDiff("n", WT(0).dStartTime, WT(Inc).dEndTime) <= Max4h30m Then
                        If DateDiff("n", dt(Inc_DT - 1).dEndTime, WT(Inc + 1).dEndTime) >= MaxTimeBetweenDT And isDonePause = False Then
                            If AddPause(dt(Inc_DT - 1).dEndTime, WT(Inc + 1).dEndTime) = False Then
                                Exit Function
                            Else
                                isDonePause = True
                                isPauseFinished = True
                                isPauseBlock = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next Inc
        End If
    '******************************************************************************************************************************
        
        If isPauseBlock = False And (isForce2Pauses = True Or isForceSinglePause = True) And DrivingPauses(0).isEnablePause = True Then
            If isAutoSave = False Then
                MsgBox "Verificar as Pausas, n�o foram geradas todas as pausas!", vbCritical + vbOKOnly
            End If
            
            hasProcessError = True
            sError = "GetDispTime: Verificar as Pausas, n�o foram geradas todas as pausas!"
        ElseIf isForce2Pauses = True And UBound(DrivingPauses) <= 1 Then
            If isAutoSave = False Then
                MsgBox "Verificar as Pausas, n�o foram geradas todas as pausas!", vbCritical + vbOKOnly
            End If
            
            hasProcessError = True
            sError = "GetDispTime: Verificar as Pausas, n�o foram geradas todas as pausas!"
        End If
    End If

    GetDispTime = MatchDispTimeWithWorkTime

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetDispTime", Err.Description
    hasProcessError = True
    sError = "GetDispTime: " & Err.Description
End Function

Private Function MatchDispTimeWithWorkTime() As Boolean
    Dim Inc, Inc_1 As Integer
    Dim isNew, BResult As Boolean
    Dim iResult As Integer
    Dim DBP_ToAdd() As DispByPeriod
    Dim dSupportDate As Date
    Dim Inc_DBP_ToAdd As Integer
    Dim sSplit() As String
    Dim dTotalTime As Double
On Error GoTo ErrHandler
 
    For Inc = 0 To Inc_DBP - 1
        For Inc_1 = 0 To Inc_WT - 1
            If Inc_1 < Inc_WT - 1 And DateDiff("n", WT(Inc_1 + 1).dStartTime, WT(Inc_1 + 1).dEndTime) <> 0 Then
                iResult = ContainsIntervalTime(WT(Inc_1).dEndTime, WT(Inc_1 + 1).dStartTime, DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                Select Case iResult
                    Case isBeforeToAfter
                        DBP(Inc).isEnabled = False
                        '*************************
                        'ERRO DE TROCO QUE TERMINA DENTRO DO HOR�RIO
                        '*************************
                        DBP(Inc).dStartTime = WT(Inc_1).dEndTime
                        DBP(Inc).dEndTime = WT(Inc_1).dEndTime
                        DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                        '*************************
                    Case isBeforeToWithin
                        DBP(Inc).dStartTime = WT(Inc_1 + 1).dStartTime
                        DBP(Inc).dEndTime = DBP(Inc).dEndTime
                        DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                        DBP(Inc).isEnabled = True
                    Case isWithinToAfter
                        DBP(Inc).dStartTime = DBP(Inc).dStartTime
                        DBP(Inc).dEndTime = WT(Inc_1).dEndTime
                        DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                        DBP(Inc).isEnabled = True
                    Case isWithin
                        If DateDiff("n", DBP(Inc).dStartTime, WT(Inc_1).dEndTime) > 0 Then
                            dSupportDate = DBP(Inc).dEndTime
                            DBP(Inc).dStartTime = DBP(Inc).dStartTime
                            DBP(Inc).dEndTime = WT(Inc_1).dEndTime
                            DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, WT(Inc_1).dEndTime)
                            DBP(Inc).isEnabled = True
                            isNew = True
                        End If
                        
                        If DateDiff("n", DBP(Inc).dEndTime, WT(Inc_1 + 1).dStartTime) > 0 Then
                            If isNew = True Then
                                'CRIAR MAIS UM INTERVALO NO FIM
                                Inc_DBP_ToAdd = Inc_DBP_ToAdd + 1
                                ReDim Preserve DBP_ToAdd(Inc_DBP_ToAdd)
                                DBP_ToAdd(Inc_DBP_ToAdd - 1).dStartTime = WT(Inc_1 + 1).dStartTime
                                DBP_ToAdd(Inc_DBP_ToAdd - 1).dEndTime = dSupportDate
                                DBP_ToAdd(Inc_DBP_ToAdd - 1).iDispTime = DateDiff("n", WT(Inc_1 + 1).dStartTime, dSupportDate)
                                DBP_ToAdd(Inc_DBP_ToAdd - 1).isEnabled = True
                            Else
                                DBP(Inc).dStartTime = WT(Inc_1 + 1).dStartTime
                                DBP(Inc).dEndTime = DBP(Inc).dEndTime
                                DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                                DBP(Inc).isEnabled = True
                            End If
                        Else
                            isNew = False
                            If DateDiff("n", WT(Inc_1 + 1).dStartTime, DBP(Inc).dEndTime) > 0 Then
                                DBP(Inc).dStartTime = WT(Inc_1 + 1).dStartTime
                                DBP(Inc).dEndTime = DBP(Inc).dEndTime
                                DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                                DBP(Inc).isEnabled = True
                            Else
                                If DateDiff("n", WT(Inc_1).dEndTime, DBP(Inc).dStartTime) = 0 And DateDiff("n", DBP(Inc).dEndTime, WT(Inc_1 + 1).dStartTime) = 0 Then
                                    DBP(Inc).iDispTime = 0
                                    DBP(Inc).isEnabled = False
                                Else
                                    DBP(Inc).isEnabled = False
                                End If
                            End If
                        End If
                    Case isOutBefore
                        If DBP(Inc).iDispTime < MaxTimeBetweenDT Then
                            If DateDiff("n", WT(Inc_1 + 1).dStartTime, DBP(Inc).dStartTime) = 0 Then
                                DBP(Inc).isEnabled = True
                            Else
                                DBP(Inc).isEnabled = False
                            End If
                        Else
                            DBP(Inc).isEnabled = True
                        End If
                    Case isOutAfter
                        If DBP(Inc).iDispTime < MaxTimeBetweenDT Then
                            If DateDiff("n", DBP(Inc).dEndTime, WT(Inc_1).dEndTime) = 0 Then
                                DBP(Inc).isEnabled = True
                            Else
                                DBP(Inc).isEnabled = False
                            End If
                        Else
                            If DBP(Inc).iDispTime <> 0 Then
                                DBP(Inc).isEnabled = True
                            End If
                        End If
                    Case isNoCalc
                        If isAutoSave = False Then
                            MsgBox "isNoCalc", vbInformation + vbOKOnly
                        End If
                        hasProcessError = True
                        sError = "MatchDispTimeWithWorkTime: Falha no calculo!"
                        Err.Raise -1, "AbstractInterval:MatchDispTimeWithWorkTime", "Falha no calculo!"
                        Exit Function
                End Select
            End If
        Next Inc_1
    Next Inc
    
    BResult = True
    
    If isNew = True Then
        For Inc = 0 To Inc_DBP_ToAdd - 1
            If addTimeInterval_DispByPeriod(DBP_ToAdd(Inc).dStartTime, DBP_ToAdd(Inc).dEndTime, False, DBP_ToAdd(Inc).isInibeDisp) = False Then
                BResult = False
            End If
            With DBP(Inc_DBP - 1)
                .iDispTime = DBP_ToAdd(Inc).iDispTime
                .isEnabled = DBP_ToAdd(Inc).isEnabled
                
                If .iDispTime < MinimoTempoDISPEmRefeicao Then
                    .isEnabled = False
                End If
            End With
            
        Next Inc
        
        ReDim DBP_ToAdd(0)
        Inc_DBP_ToAdd = 0
    End If
    
    'Verify iDispTime
    For Inc = 0 To Inc_DBP - 1
        If DBP(Inc).iDispTime < MinimoTempoDISPEmRefeicao Then
            DBP(Inc).isEnabled = False
        End If
    Next Inc
    
    If OrderDBP() = False Then
        'NADA
    End If

    For Inc = 0 To Inc_DBP - 1
        If DBP(Inc).isEnabled = True Then
            If DBP(Inc).iDispTime > 0 Then
                dTotalTime = dTotalTime + DBP(Inc).iDispTime
            Else
                DBP(Inc).isEnabled = False
            End If
        End If
    Next Inc

    sSplit() = Split(CStr(dTotalTime), ",")
    If UBound(sSplit()) > 0 Then
        If IsDate(ElapsedTimeDouble(sSplit())) = True Then
            dFullDispTimeAvailable = CDate(ElapsedTimeDouble(sSplit()))
        Else
            dFullDispTimeAvailable = CDate(0)
        End If
    Else
        If sSplit(0) <= 86400 Then
            dFullDispTimeAvailable = CDate(ElapsedTimeInt(sSplit(0)))
        Else
            dFullDispTimeAvailable = CDate(0)
        End If
    End If
    
    MatchDispTimeWithWorkTime = BResult

Exit Function
ErrHandler:
    If Err.Number = 6 Then 'overflow horario'
        MsgBox "Verifique hor�rios. N�o � possivel calcular!", vbInformation
    Else
        Err.Raise Err.Number, "AbstractInterval:MatchDispTimeWithWorkTime", Err.Description
        hasProcessError = True
        sError = "MatchDispTimeWithWorkTime: " & Err.Description
    End If
End Function

Private Function MatchPAUSETimeWithWorkTime(ByRef dStartDate As Date, ByRef dEndDate As Date) As Boolean
    Dim Inc, Inc_1 As Integer
    Dim isNew, BResult As Boolean
    Dim iResult As Integer
    Dim dSupportDate As Date
On Error GoTo ErrHandler

    For Inc_1 = 0 To Inc_WT - 1
        If Inc_1 < Inc_WT - 1 And DateDiff("n", WT(Inc_1 + 1).dStartTime, WT(Inc_1 + 1).dEndTime) <> 0 Then
            iResult = ContainsIntervalTime(dStartDate, dEndDate, WT(Inc_1).dEndTime, WT(Inc_1 + 1).dStartTime)
            Select Case iResult
                Case isBeforeToAfter
                    dEndDate = WT(Inc_1).dEndTime
                    MatchPAUSETimeWithWorkTime = False
                Case isBeforeToWithin
                    dEndDate = WT(Inc_1).dEndTime
                    MatchPAUSETimeWithWorkTime = False
                Case isWithinToAfter
                    dStartDate = WT(Inc_1 + 1).dStartTime
                    MatchPAUSETimeWithWorkTime = False
                Case isWithin
                    dEndDate = dStartDate
                    MatchPAUSETimeWithWorkTime = False
                Case isOutBefore
                    MatchPAUSETimeWithWorkTime = True
                Case isOutAfter
                    MatchPAUSETimeWithWorkTime = True
                Case isNoCalc
                    MsgBox "isNoCalc", vbInformation + vbOKOnly
                    Err.Raise -1, "AbstractInterval:MatchDispTimeWithWorkTime", "Falha no calculo!"
                    Exit Function
            End Select
        End If
    Next Inc_1

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:MatchPAUSETimeWithWorkTime", Err.Description
    hasProcessError = True
    sError = "MatchPAUSETimeWithWorkTime: " & Err.Description
End Function


Private Function OrderDBP() As Boolean
    Dim Inc As Integer
    Dim IncAux As Integer
    Dim DBP_Aux As DispByPeriod
    On Error GoTo ErrHandler

    '******************************************************************************************************
    'REORGANIZAR
    '******************************************************************************************************
    For Inc = 1 To Inc_DBP - 1
        For IncAux = Inc + 1 To Inc_DBP
            If DateTime.DateDiff("n", DBP(Inc).dStartTime, DBP(IncAux).dStartTime) < 0 And DBP(IncAux).dStartTime <> DBP(IncAux).dEndTime Then
                
                With DBP_Aux
                    .dStartTime = DBP(IncAux).dStartTime
                    .dEndTime = DBP(IncAux).dEndTime
                    .iDispTime = DBP(IncAux).iDispTime
                    .isEnabled = DBP(IncAux).isEnabled
                End With
                
                DBP(IncAux) = DBP(Inc)
                DBP(Inc) = DBP_Aux
            End If
        Next IncAux
    Next Inc
    '******************************************************************************************************
    
    OrderDBP = True
    
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "OrderDBP", Err.Description
    OrderDBP = False
    hasProcessError = True
    sError = "OrderDBP: " & Err.Description
End Function

Public Function AlgorithTimeMatchCalculation() As Boolean
    Dim Inc, Inc_1 As Integer
    Dim bBeginProcess As Boolean
    Dim iResult As Integer
On Error GoTo ErrHandler

    For Inc = 0 To Inc_DP - 1
        For Inc_1 = 0 To Inc_WT - 1
            iResult = ContainsIntervalTime(WT(Inc_1).dStartTime, WT(Inc_1).dEndTime, DP(Inc).dStartTime, DP(Inc).dEndTime)
            Select Case iResult
                Case isBeforeToAfter
                    bBeginProcess = True
                    MsgBox "isBeforeToAfter", vbInformation + vbOKOnly
                Case isBeforeToWithin
                    bBeginProcess = True
                    MsgBox "isBeforeToWithin", vbInformation + vbOKOnly
                Case isWithinToAfter
                    bBeginProcess = True
                    MsgBox "isWithinToAfter", vbInformation + vbOKOnly
                Case isWithin
                    bBeginProcess = True
                    MsgBox "isWithin", vbInformation + vbOKOnly
                Case isOutBefore
                    bBeginProcess = False
                    MsgBox "isOutBefore", vbInformation + vbOKOnly
                Case isOutAfter
                    bBeginProcess = False
                    MsgBox "isOutAfter", vbInformation + vbOKOnly
                Case isNoCalc
                    bBeginProcess = False
                    MsgBox "isNoCalc", vbInformation + vbOKOnly
                    Err.Raise -1, "AbstractInterval:AlgorithTimeMatch", "Falha no calculo!"
                    Exit Function
            End Select
            
            If bBeginProcess = True Then
                If LaborTimes(DP(Inc).dStartTime, DP(Inc).dEndTime, WT(Inc_1).dStartTime, WT(Inc_1).dEndTime, iResult) = False Then
                    Exit Function
                End If
            End If
        Next
    Next Inc

AlgorithTimeMatchCalculation = True

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:AlgorithTimeMatchCalculation", Err.Description
    hasProcessError = True
    sError = "AlgorithTimeMatchCalculation: " & Err.Description
End Function

Private Function FinalyzeDispTimeCalculation(ByVal dExtraTime As Date) As Boolean
    Dim Inc As Integer
    Dim dAvailableTime, dSupport, dResult As Double
    Dim sDispTime As String
    Dim sSplit() As String
On Error GoTo ErrHandler
    
    dAvailableTime = dExtraTime * 60 * 24
    
    For Inc = 0 To Inc_DBP - 1
        If DBP(Inc).isEnabled = True Then
            If dAvailableTime > 0 Then
                Inc_CDT = Inc_CDT + 1
                ReDim Preserve CDT(Inc_CDT)
                If dAvailableTime >= DBP(Inc).iDispTime Then
                    dAvailableTime = dAvailableTime - DBP(Inc).iDispTime
                    CDT(Inc_CDT - 1).dDispStartTime = DBP(Inc).dStartTime
                    CDT(Inc_CDT - 1).dDispEndTime = DBP(Inc).dEndTime
                    CDT(Inc_CDT - 1).iDispTimeMinutes = DBP(Inc).iDispTime
                    
                    sSplit() = Split(CStr(DBP(Inc).iDispTime), ",")
                    If UBound(sSplit()) > 0 Then
                        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
                    Else
                        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
                    End If

                    CDT(Inc_CDT - 1).iDispTimeHours = CDate(sDispTime)
                Else
                    DBP(Inc).dEndTime = DateAdd("n", dAvailableTime, DBP(Inc).dStartTime)
                    DBP(Inc).iDispTime = DateDiff("n", DBP(Inc).dStartTime, DBP(Inc).dEndTime)
                    CDT(Inc_CDT - 1).dDispStartTime = DBP(Inc).dStartTime
                    CDT(Inc_CDT - 1).dDispEndTime = DBP(Inc).dEndTime
                    CDT(Inc_CDT - 1).iDispTimeMinutes = DBP(Inc).iDispTime
                    
                    sSplit() = Split(CStr(DBP(Inc).iDispTime), ",")
                    If UBound(sSplit()) > 0 Then
                        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
                    Else
                        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
                    End If

                    CDT(Inc_CDT - 1).iDispTimeHours = CDate(sDispTime)
                    
                    dAvailableTime = 0
                End If
            Else
                DBP(Inc).isEnabled = False
            End If
        End If
    Next Inc

FinalyzeDispTimeCalculation = True

Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:FinalyzeDispTimeCalculation", Err.Description
    hasProcessError = True
    sError = "FinalyzeDispTimeCalculation: " & Err.Description
End Function

Private Function LaborTimes(ByVal dStart_DPTime As Date, ByVal dEnd_DPTime As Date, ByVal dStart_WTTime As Date, ByVal dEnd_WTTime As Date, ByVal iIntervalType As Integer) As Boolean
    Dim Inc As Integer
On Error GoTo ErrHandler

    Inc = Inc + 1
    
    ReDim Preserve AWTDT(Inc)
    
    With AWTDT(Inc - 1)
        .dDispStartTime = dStart_DPTime
        .dDispEndTime = dEnd_DPTime
        .dWorkStartTime = dStart_WTTime
        .dWorkEndTime = dEnd_WTTime
        .iWorkedTimePeriod = iIntervalType
        
        Select Case iIntervalType
            Case isBeforeToAfter
                .iWorkedTime = DateDiff("n", .dDispStartTime, .dDispEndTime)
                TimeToNextPeriod = DateDiff("n", .dDispEndTime, .dWorkEndTime)
            Case isBeforeToWithin
                .iWorkedTime = DateDiff("n", .dDispStartTime, .dWorkEndTime)
                TimeToNextPeriod = 0
            Case isWithinToAfter
                .iWorkedTime = DateDiff("n", .dWorkStartTime, .dDispEndTime)
                TimeToNextPeriod = DateDiff("n", .dDispEndTime, .dWorkEndTime)
            Case isWithin
                .iWorkedTime = DateDiff("n", .dWorkStartTime, .dWorkEndTime)
                TimeToNextPeriod = 0
        End Select
        
        If iLastiIntervalType > iIntervalType Then
            .iWorkedTime = .iWorkedTime + TimeToNextPeriod
        End If
        
        iLastiIntervalType = iIntervalType
    End With
    
    LaborTimes = True
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:GetOverTimeTimeAfterGivenEnd", Err.Description
    hasProcessError = True
    sError = "LaborTimes: " & Err.Description
End Function

Private Function CalculateDispTimeByPeriod() As Boolean
    Dim Inc, Inc_1, iIntervalType As Integer
    Dim sDispTime As String
    Dim sSplit() As String
On Error GoTo ErrHandler
    
    For Inc = 0 To Inc_DP - 1
        For Inc_1 = 0 To Inc_CDT - 1
            iIntervalType = ContainsIntervalTime(CDT(Inc_1).dDispStartTime, CDT(Inc_1).dDispEndTime, DP(Inc).dStartTime, DP(Inc).dEndTime)
            Select Case iIntervalType
                Case isBeforeToAfter
                    DP(Inc).iDispTimeMinutes = DP(Inc).iDispTimeMinutes + DateDiff("n", DP(Inc).dStartTime, DP(Inc).dEndTime)
                    DP(Inc).iDispTimeHours = DP(Inc).iDispTimeMinutes * 60
                    sSplit() = Split(CStr(DP(Inc).iDispTimeMinutes), ",")
                    If UBound(sSplit()) > 0 Then
                        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
                    Else
                        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
                    End If
                    
                    DP(Inc).iDispTimeHours = CDate(sDispTime)
                Case isBeforeToWithin
                    DP(Inc).iDispTimeMinutes = DP(Inc).iDispTimeMinutes + DateDiff("n", DP(Inc).dStartTime, CDT(Inc_1).dDispEndTime)
                    DP(Inc).iDispTimeHours = DP(Inc).iDispTimeMinutes * 60
                    sSplit() = Split(CStr(DP(Inc).iDispTimeMinutes), ",")
                    If UBound(sSplit()) > 0 Then
                        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
                    Else
                        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
                    End If
                    
                    DP(Inc).iDispTimeHours = CDate(sDispTime)
                Case isWithinToAfter
                    DP(Inc).iDispTimeMinutes = DP(Inc).iDispTimeMinutes + DateDiff("n", CDT(Inc_1).dDispStartTime, DP(Inc).dEndTime)
                    DP(Inc).iDispTimeHours = DP(Inc).iDispTimeMinutes * 60
                    sSplit() = Split(CStr(DP(Inc).iDispTimeMinutes), ",")
                    If UBound(sSplit()) > 0 Then
                        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
                    Else
                        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
                    End If
                    
                    DP(Inc).iDispTimeHours = CDate(sDispTime)
                Case isWithin
                    DP(Inc).iDispTimeMinutes = DP(Inc).iDispTimeMinutes + DateDiff("n", CDT(Inc_1).dDispStartTime, CDT(Inc_1).dDispEndTime)
                    sSplit() = Split(CStr(DP(Inc).iDispTimeMinutes), ",")
                    If UBound(sSplit()) > 0 Then
                        sDispTime = CDate(ElapsedTimeDouble(sSplit()))
                    Else
                        sDispTime = CDate(ElapsedTimeInt(sSplit(0)))
                    End If
                    
                    DP(Inc).iDispTimeHours = CDate(sDispTime)
            End Select
        Next Inc_1
    Next Inc

    CalculateDispTimeByPeriod = True
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, "AbstractInterval:CalculateDispTimeByPeriod", Err.Description
    hasProcessError = True
    sError = "CalculateDispTimeByPeriod: " & Err.Description
End Function

Private Function ElapsedTimeDouble(ByRef sSplit() As String) As String
    Dim hours As Integer
    Dim minutes As Integer
    Dim seconds As Integer
    Dim sSplitDim As Integer
    On Error GoTo ErrHandler
    
    hours = sSplit(0)
    minutes = (CDbl("0," & sSplit(1))) * 60
    
    If minutes = 60 Then
        Do While minutes >= 60
            hours = hours + 1
            minutes = minutes - 60
        Loop
    End If
    
    ElapsedTimeDouble = CStr(hours) & ":" & minutes & ":" & seconds
    Exit Function
ErrHandler:
    If isAutoSave = False Then
        MsgBox "Erro na convers�o do Tempo de Ag�nte �nico!", vbCritical + vbOKOnly, "ElapsedTime"
    End If
    ElapsedTimeDouble = "00:00:00"
    hasProcessError = True
    sError = "ElapsedTimeDouble: Erro na convers�o do Tempo de Ag�nte �nico!"
End Function

Private Function ElapsedTimeInt(ByVal seconds As Integer) As String
    Dim sTime As String
    On Error GoTo ErrHandler
    
    seconds = seconds Mod 3600
    sTime = sTime & Format$(seconds \ 60, "00:")
    sTime = sTime & Format$(seconds Mod 60, "00")
    ElapsedTimeInt = sTime
    Exit Function
ErrHandler:
    If isAutoSave = False Then
        MsgBox "Erro na convers�o do Tempo de Ag�nte �nico!", vbCritical + vbOKOnly, "ElapsedTime"
    End If
    ElapsedTimeInt = "00:00:00"
    hasProcessError = True
    sError = "ElapsedTimeInt: Erro na convers�o do Tempo de Ag�nte �nico!"
End Function
'****************************************************************************************************************


'****************************************************************************************************************
'DRIVING PAUSE
'WHEN ( [CARREIRA OR ALUGUER] >= 50 KMS OR OCASIONAL
'   AND DRIVING TIME >= 4H30M
'       SELECTION OPTIONS
'       - FORCES 1� STOP 15 MIN AND 2� STOP 20 MIN
'           - STOPS MUST BE UNTIL 4H30M, CANNOT BE GIVER AFTER 4H30MIN
'           - IF STOP CANNOT BE APPLIED ERROR MUST BE GIVEN
'               - CANNOT SET THE TWO PAUSES ("N�o foi poss�vel atribu�r as pausas!")
'               - CANNOT SET THE SECOND PAUSE ("N�o foi poss�vel atribu�r a 2� pausa!")
'       - FORCES SINGLE STOP 45 MIN
'           - STOPS MUST BE UNTIL 4H30M, CANNOT BE GIVER AFTER 4H30MIN
'           - IF STOP CANNOT BE APPLIED ERROR MUST BE GIVEN
'               - CANNOT SET THE PAUSE ("N�o foi poss�vel atribu�r a pausa!")
'       - DOES NOTHING
'           - DO NOTHING SELECTED
'****************************************************************************************************************

Public Function AddPauses(ByVal isCarreiraPlus49Kms As Boolean, ByVal isAluguerPlus49Kms As Boolean, _
                    ByVal isOcasional As Boolean, ByVal isClear As Boolean) As Boolean
                    
    Dim Inc As Long
    On Error GoTo ErrHandler

    If isClear = True Then
        Inc_Pauses = 0
        Inc_Pause = 0
        ReDim DrivingPauses(Inc_Pauses)
        ReDim DrivingPause(Inc_Pause)
    End If

    Inc_Pauses = Inc_Pauses + 1
    ReDim Preserve DrivingPauses(Inc_Pauses)
    
    With DrivingPauses(Inc_Pauses - 1)
        .isAluguerPlus49Kms = isAluguerPlus49Kms
        .isCarreiraPlus49Kms = isCarreiraPlus49Kms
        .isOcasional = isOcasional
        .iTotalDrivingTime = GetDriveTime()
        
        If .isAluguerPlus49Kms = True Or .isCarreiraPlus49Kms = True Or .isOcasional = True Then
            .isEnablePause = True
        Else
            .isEnablePause = False
        End If
    End With
    
    AddPauses = True

    Exit Function
ErrHandler:
    If isAutoSave = False Then
        MsgBox "Erro ao atribuir a pausa!", vbCritical + vbOKOnly, "AddPauses"
    End If
    hasProcessError = True
    sError = "AddPauses: Erro ao atribuir a pausa!"
End Function

Public Function AddPause(ByVal dStartTime As Date, ByVal dEndTime As Date) As Boolean
        Dim Inc As Long
        Dim iDefaultValue As Integer
    On Error GoTo ErrHandler
    
    If DrivingPauses(0).isEnablePause = True And Inc_Pause <= 1 Then
        If isForce2Pauses = True Then
            If Inc_Pause = 0 Then
                iDefaultValue = 15
            Else
                iDefaultValue = 30
            End If
        ElseIf isForceSinglePause Then
            iDefaultValue = 45
        End If
        
        For Inc = 0 To Inc_Pause - 1
            If DateDiff("n", dStartTime, DrivingPause(Inc).dStartTime) = 0 Then
                AddPause = True
                Exit Function
            End If
        Next Inc
        
        If DateDiff("n", dStartTime, dEndTime) >= iDefaultValue Then
            Inc_Pause = Inc_Pause + 1
            ReDim Preserve DrivingPause(Inc_Pause)
            
            With DrivingPause(Inc_Pause - 1)
                .dStartTime = dStartTime
                .dEndTime = DateAdd("n", iDefaultValue, dStartTime)
                .iTotalPauseTime = DateDiff("n", .dStartTime, .dEndTime)
            End With
        End If
    End If
    
    AddPause = True

    Exit Function
ErrHandler:
    If isAutoSave = False Then
        MsgBox "Erro ao atribuir a pausa!", vbCritical + vbOKOnly, "AddPause"
    End If
    hasProcessError = True
    sError = "AddPause: Erro ao atribuir a pausa!"
End Function
