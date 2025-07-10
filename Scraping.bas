Attribute VB_Name = "Module5"
' ScrapingModule�̐擪�ɒǉ�
Public Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)

'--- �X�N���C�s���O���W���[���̍œK�� ---
' �L���b�V���̗L�������ݒ�i�Z�b�V�������ł̂ݗL���j
Private CacheEnabled As Boolean
Private MaxCacheSize As Long
Private TransitCache As Object ' Dictionary�^

' �L���b�V���̏����� - �œK����
Private Sub InitializeCache()
    If TransitCache Is Nothing Then
        Set TransitCache = CreateObject("Scripting.Dictionary")
        CacheEnabled = True
        MaxCacheSize = 6000  '
    End If
End Sub

' ���ʂ��L���b�V�� - �œK����
Private Sub CacheTransitResult(fromAddr As String, toAddr As String, result As Variant)
    If Not CacheEnabled Then Exit Sub
    
    InitializeCache
    Dim key As String
    key = MakeCacheKey(fromAddr, toAddr)
    
    ' �L���b�V���T�C�Y����
    If TransitCache.COUNT >= MaxCacheSize Then
        ' �ł��Â��G���g�����폜�i�ȈՎ����j
        Dim firstKey As Variant
        firstKey = TransitCache.keys()(0)
        TransitCache.Remove firstKey
    End If
    
    If Not TransitCache.Exists(key) Then
        TransitCache.Add key, result
    End If
End Sub

' �X�N���C�s���O�œK���ҋ@���Ԏ擾
Public Function GetOptimizedWaitTime(minMs As Integer, maxMs As Integer) As Integer
    ' �ʏ��菭���Z���ҋ@���Ԃ�Ԃ�
    GetOptimizedWaitTime = minMs + Int(Rnd() * (maxMs - minMs + 1))
End Function


' �L���b�V���L�[�쐬�i�o���n+�ړI�n�j
Private Function MakeCacheKey(fromAddr As String, toAddr As String) As String
    MakeCacheKey = fromAddr & "��" & toAddr
End Function


' �L���b�V���m�F
Private Function GetCachedTransitResult(fromAddr As String, toAddr As String) As Variant
    InitializeCache
    Dim key As String
    key = MakeCacheKey(fromAddr, toAddr)
    
    ' �ʏ�����̃L���b�V�����m�F
    If TransitCache.Exists(key) Then
        GetCachedTransitResult = TransitCache(key)
        Exit Function
    End If
    
    ' �t�����̃L���b�V�����m�F
    Dim reverseKey As String
    reverseKey = GetReverseCacheKey(fromAddr, toAddr)
    
    If TransitCache.Exists(reverseKey) Then
        Dim reverseResult As Variant
        reverseResult = TransitCache(reverseKey)
        
        ' �t�����̌��ʂ𒲐����ĕԂ�
        Dim adjustedResult As Variant
        adjustedResult = AdjustReverseTransitResult(reverseResult)
        
        If Not IsNull(adjustedResult) Then
            ' �����������ʂ�ʏ�����̃L�[�ł��L���b�V���ɕۑ�
            If Not TransitCache.Exists(key) Then
                TransitCache.Add key, adjustedResult
            End If
            
            GetCachedTransitResult = adjustedResult
            Exit Function
        End If
    End If
    
    ' �L���b�V���Ɍ�����Ȃ��ꍇ��Null
    GetCachedTransitResult = Null
End Function

' �t�����̃L�[���쐬����֐��i�V�K�쐬�j
Private Function GetReverseCacheKey(fromAddr As String, toAddr As String) As String
    GetReverseCacheKey = toAddr & "��" & fromAddr
End Function

' �t�����̌��ʂ𒲐�����֐��i�V�K�쐬�j
Private Function AdjustReverseTransitResult(reverseResult As Variant) As Variant
    Dim adjustedResult(2) As Variant
    
    If IsArray(reverseResult) Then
        ' ���Ԃ͂��̂܂܎g�p�i�����łقړ����Ɖ���j
        adjustedResult(0) = reverseResult(0)
        
        ' �w���X�g���t���ɕϊ�
        Dim originalStations As String
        originalStations = reverseResult(1)
        
        If originalStations <> "" Then
            Dim stationArray As Variant
            stationArray = Split(originalStations, " �� ")
            
            ' �z����t���ɂ���
            Dim reversedStations As String
            reversedStations = ""
            
            Dim i As Integer
            For i = UBound(stationArray) To 0 Step -1
                If reversedStations = "" Then
                    reversedStations = stationArray(i)
                Else
                    reversedStations = reversedStations & " �� " & stationArray(i)
                End If
            Next i
            
            adjustedResult(1) = reversedStations
        Else
            adjustedResult(1) = ""
        End If
        
        ' ���ԃe�L�X�g�͂��̂܂܎g�p
        adjustedResult(2) = reverseResult(2)
    Else
        ' �t�����̌��ʂ������ȏꍇ��Null��Ԃ�
        AdjustReverseTransitResult = Null
        Exit Function
    End If
    
    AdjustReverseTransitResult = adjustedResult
End Function


' ���ǔŗv�f�ҋ@�֐�
Private Function WaitFindElementDynamic(ByVal drv As SeleniumVBA.WebDriver, _
                                       ByVal cssSel As String, _
                                       ByVal baseTimeout As Long) As SeleniumVBA.WebElement
    ' �����_���Ȓǉ�����
    Dim extraTime As Long
    extraTime = Int(Rnd() * 3)
    
    Dim tEnd As Date
    tEnd = DateAdd("s", baseTimeout + extraTime, Now)
    
    Dim el As SeleniumVBA.WebElement
    
    Do
        On Error Resume Next
        Set el = drv.FindElement(By.cssSelector, cssSel)
        On Error GoTo 0
        
        If Not el Is Nothing Then Exit Do
        DoEvents
        ' �Z���Ԋu�ł̃|�[�����O������邽�߂ɏ����ȑҋ@������
        drv.Wait 100
    Loop While Now < tEnd
    
    If Not el Is Nothing Then
        Set WaitFindElementDynamic = el
    Else
        Set WaitFindElementDynamic = Nothing
    End If
End Function


' ���S�ȃX���[�v�֐��iDoEvents�Ăяo���̍팸�j
Public Sub SafeSleep(ByVal ms As Long)
    ' DoEvents�Ȃ���WindowsAPI�o�R�̃X���[�v�����s
    SleepWinAPI ms
End Sub

' ���S�ȗv�f�ҋ@�֐��i�^�C���A�E�g�@�\�����j
Public Function SafeWaitFindElement(ByVal driver As SeleniumVBA.WebDriver, _
                                    ByVal cssSelector As String, _
                                    ByVal timeoutSec As Long) As SeleniumVBA.WebElement
    Dim EndTime As Date
    EndTime = DateAdd("s", timeoutSec, Now)
    
    Dim element As SeleniumVBA.WebElement
    Dim doEventsCounter As Long: doEventsCounter = 0
    
    Do
        On Error Resume Next
        Set element = driver.FindElement(By.cssSelector, cssSelector)
        On Error GoTo 0
        
        If Not element Is Nothing Then Exit Do
        
        ' DoEvents�Ăяo���𐧌��i10���1��j
        doEventsCounter = doEventsCounter + 1
        If doEventsCounter Mod 10 = 0 Then DoEvents
        
        ' ���Z���ҋ@
        SleepWinAPI 100
    Loop While Now < EndTime
    
    Set SafeWaitFindElement = element
End Function

' �C���� GetTransitInfo�֐�
Public Function GetTransitInfo(ByVal driver As SeleniumVBA.WebDriver, _
                              ByVal fromAddress As String, _
                              ByVal toAddress As String) As Variant
    Dim result(2) As Variant  ' ���ʊi�[�z��
    result(0) = 999           ' ���v���ԁi���j
    result(1) = ""            ' �w���X�g
    result(2) = ""            ' ���v���ԃe�L�X�g
    
    ' �o���n�ƖړI�n�������ꍇ��0��Ԃ�
    If fromAddress = toAddress Then
        result(0) = 0
        GetTransitInfo = result
        Exit Function
    End If
    
    ' �L���b�V���`�F�b�N
    Dim cachedResult As Variant
    cachedResult = GetCachedTransitResult(fromAddress, toAddress)
    If Not IsNull(cachedResult) Then
        GetTransitInfo = cachedResult
        Exit Function
    End If
    
    ' �ő���s���Ԃ̐ݒ�
    Dim maxExecutionTime As Long: maxExecutionTime = 30 ' �b
    Dim StartTime As Date: StartTime = Now
    
    On Error GoTo ErrorHandler
    
    ' Yahoo�H�����ɃA�N�Z�X
    driver.NavigateTo "https://transit.yahoo.co.jp/"
    SafeSleep GetOptimizedWaitTime(600, 1000)
    
    ' ���s���Ԓ��߃`�F�b�N
    If DateDiff("s", StartTime, Now) > maxExecutionTime Then GoTo TimeoutHandler
    
    ' �����t�H�[������
    Dim fromInput As SeleniumVBA.WebElement
    Set fromInput = SafeWaitFindElement(driver, "input[name='from']", 5)
    If fromInput Is Nothing Then GoTo ErrorHandler
    
    fromInput.Clear
    fromInput.SendKeys fromAddress
    
    Dim toInput As SeleniumVBA.WebElement
    Set toInput = driver.FindElement(By.cssSelector, "input[name='to']")
    toInput.Clear
    toInput.SendKeys toAddress
    
    SafeSleep 300
    
    ' ���s���Ԓ��߃`�F�b�N
    If DateDiff("s", StartTime, Now) > maxExecutionTime Then GoTo TimeoutHandler
    
    ' �����ݒ�i9:00�j
    Dim selHH As SeleniumVBA.WebElement
    Set selHH = driver.FindElement(By.cssSelector, "select[name='hh']")
    If Not selHH Is Nothing Then
        selHH.Click
        SafeSleep 300
        
        Dim optHH As SeleniumVBA.WebElement
        Set optHH = selHH.FindElement(By.cssSelector, "option[value='09']")
        If Not optHH Is Nothing Then optHH.Click
    End If
    
    ' ���ݒ�i00���j
    SafeSleep 300
    Dim selMM As SeleniumVBA.WebElement
    Set selMM = driver.FindElement(By.cssSelector, "select[name='mm']")
    If Not selMM Is Nothing Then
        selMM.Click
        SafeSleep 300
        
        Dim optMM As SeleniumVBA.WebElement
        Set optMM = selMM.FindElement(By.cssSelector, "option[value='00']")
        If Not optMM Is Nothing Then optMM.Click
    End If
    
    ' �����{�^���N���b�N
    SafeSleep 300
    driver.FindElement(By.ID, "searchModuleSubmit").Click
    
    ' ���s���Ԓ��߃`�F�b�N
    If DateDiff("s", StartTime, Now) > maxExecutionTime Then GoTo TimeoutHandler
    
    ' ���ʎ擾
    Dim summaryBlock As SeleniumVBA.WebElement
    Set summaryBlock = SafeWaitFindElement(driver, "div#route01 div.routeSummary ul.summary", 7)
    If summaryBlock Is Nothing Then GoTo ErrorHandler
    
    Dim timeEl As SeleniumVBA.WebElement
    On Error Resume Next
    Set timeEl = summaryBlock.FindElement(By.cssSelector, "li.time")
    On Error GoTo 0
    
    If timeEl Is Nothing Then GoTo ErrorHandler
    
    ' ���ԃe�L�X�g�擾�E�ϊ�
    Dim rawTimeText As String: rawTimeText = timeEl.GetText
    Dim totalTimeText As String: totalTimeText = ExtractTimeOnly(rawTimeText)
    result(0) = ConvertTimeTextToMinutes(totalTimeText)
    result(2) = totalTimeText
    
    ' �w���ꗗ�擾
    Dim detailBlock As SeleniumVBA.WebElement
    Set detailBlock = SafeWaitFindElement(driver, "div#route01 div.routeDetail", 7)
    If detailBlock Is Nothing Then GoTo ErrorHandler
    
    Dim stationEls As SeleniumVBA.WebElements
    Set stationEls = detailBlock.FindElements(By.cssSelector, "div.station dl dt a")
    
    Dim stationList As String: stationList = ""
    If Not stationEls Is Nothing Then
        Dim st As SeleniumVBA.WebElement
        For Each st In stationEls
            Dim stName As String: stName = st.GetText
            If stationList = "" Then
                stationList = stName
            Else
                stationList = stationList & " �� " & stName
            End If
        Next
    End If
    
    result(1) = stationList
    CacheTransitResult fromAddress, toAddress, result
    
    GetTransitInfo = result
    Exit Function
    
TimeoutHandler:
    Call WriteLog("GetTransitInfo: �^�C���A�E�g - " & fromAddress & "��" & toAddress, "GetTransitInfo")
    On Error Resume Next
    driver.ExecuteScript "window.stop();"
    driver.NavigateTo "about:blank"
    SafeSleep 1000
    On Error GoTo 0
    GetTransitInfo = result
    Exit Function
    
ErrorHandler:
    Call WriteLog("GetTransitInfo: �G���[ - " & fromAddress & "��" & toAddress, "GetTransitInfo")
    GetTransitInfo = result
End Function

' GetYahooTransit_UseOptionClick_WithDriver�֐�
Public Function GetYahooTransit_UseOptionClick_WithDriver( _
    ByVal driver As SeleniumVBA.WebDriver, _
    ByVal baseAddress As String, _
    ByVal dayToSet As String, _
    ByVal checkFlg As Integer, _
    ByVal grpTime As Integer, _
    Optional ByVal storeAddress As String = "", _
    Optional ByVal excludedStations As String = "" _
) As Boolean

    Dim loopMAX             As Integer
    Dim foundIndex          As Integer
    Dim i                   As Integer
    Dim processedCount      As Integer
    Dim maxExecutionTime    As Long
    Dim StartTime           As Date
    
    ' ���ۑ��p�ϐ�
    Dim depStation          As String, arrStation As String, timeText As String
    Dim storeDepStation     As String, storeArrStation As String, storeTimeText As String
    
    On Error GoTo Finalize
    
    ' �ő���s���Ԑݒ�
    maxExecutionTime = 120 ' 2��
    StartTime = Now
    
    ' �V�[�g2����X�e�[�V�������擾
    With Sheets(COMMON_SETTING7)
        loopMAX = WorksheetFunction.CountIf(.Range("E4:E43"), "*")
    End With
    
    ' ������
    foundIndex = -1
    processedCount = 0
    depStation = "": arrStation = "": timeText = ""
    storeDepStation = "": storeArrStation = "": storeTimeText = ""
    
    ' ���O�X�e�[�V�����z��̍쐬
    Dim excludedArray() As String
    Dim excludedCount As Integer
    excludedCount = 0
    
    If Len(Trim(excludedStations)) > 0 Then
        excludedArray = Split(excludedStations, ",")
        excludedCount = UBound(excludedArray) + 1
    End If
    
    ' A�񂪋�̃X�e�[�V�����݂̂�Ώۂɂ���
    Dim targetStations() As Integer
    ReDim targetStations(loopMAX - 1)
    Dim targetCount As Integer: targetCount = 0
    
    For i = 0 To loopMAX - 1
        ' �^�C���A�E�g�`�F�b�N
        If DateDiff("s", StartTime, Now) > maxExecutionTime Then
            GoTo TimeoutHandler
        End If
        
        If Sheets(COMMON_SETTING7).Range("A" & 4 + i).Value = "" Then
            Dim stNum As String
            stNum = CStr(Sheets(COMMON_SETTING7).Range("E" & 4 + i).Value)
            
            ' ���O�X�e�[�V�����`�F�b�N
            Dim isExcluded As Boolean
            isExcluded = False
            
            If excludedCount > 0 Then
                Dim e As Integer
                For e = 0 To excludedCount - 1
                    If Trim(stNum) = Trim(excludedArray(e)) Then
                        isExcluded = True
                        Exit For
                    End If
                Next e
            End If
            
            If Not isExcluded Then
                targetStations(targetCount) = i
                targetCount = targetCount + 1
            End If
        End If
    Next i
    
    If targetCount = 0 Then
        GetYahooTransit_UseOptionClick_WithDriver = False
        GoTo Finalize
    End If
    
    ReDim Preserve targetStations(targetCount - 1)
    
    ' �e�X�e�[�V����������
    For i = 0 To targetCount - 1
        ' �^�C���A�E�g�`�F�b�N
        If DateDiff("s", StartTime, Now) > maxExecutionTime Then
            GoTo TimeoutHandler
        End If
        
        Dim idx As Integer: idx = targetStations(i)
        
        ' ���N�G�X�g�ԑҋ@
        If processedCount > 0 Then SafeSleep GetOptimizedWaitTime(1000, 2000)
        
        Dim stAddr As String: stAddr = Sheets(COMMON_SETTING7).Range("G" & 4 + idx).Value
        If stAddr = "" Then GoTo NextStation
        
        ' A�񂪋󂩂��Ċm�F
        If Sheets(COMMON_SETTING7).Range("A" & 4 + idx).Value <> "" Then GoTo NextStation
        
        Dim toBaseMin  As Integer: toBaseMin = 999
        Dim toStoreMin As Integer: toStoreMin = 999
        Dim Dep        As String, Arr As String, tText As String
        Dim sDep       As String, sArr As String, sText As String
        
        ' ���򏈗�
        Dim isValid As Boolean: isValid = False
        
        If checkFlg = 1 Then
            ' ��Z������̂ݔ���
            If baseAddress = stAddr Then
                toBaseMin = 0
                isValid = True
            Else
                Dim r1 As Variant: r1 = GetTransitInfo(driver, baseAddress, stAddr)
                If IsArray(r1) Then
                    toBaseMin = r1(0)
                    ParseStations r1(1), Dep, Arr
                    tText = r1(2)
                    If toBaseMin <= grpTime Then
                        isValid = True
                    End If
                End If
            End If
            
        ElseIf checkFlg = 2 Then
            ' ��Z�����X�e�[�V���� & �X�e�[�V�������X�� ��������
            Dim baseOK As Boolean: baseOK = False
            Dim storeOK As Boolean: storeOK = False
            
            If baseAddress = stAddr Then
                toBaseMin = 0
                baseOK = True
            Else
                Dim r2 As Variant: r2 = GetTransitInfo(driver, baseAddress, stAddr)
                If IsArray(r2) Then
                    toBaseMin = r2(0)
                    ParseStations r2(1), Dep, Arr
                    tText = r2(2)
                    If toBaseMin <= grpTime Then
                        baseOK = True
                    End If
                End If
            End If
            
            ' �x�[�X�����������Ă��Ȃ��Ȃ�X�L�b�v
            If Not baseOK Then GoTo NextStation
            
            ' �^�C���A�E�g�`�F�b�N
            If DateDiff("s", StartTime, Now) > maxExecutionTime Then
                GoTo TimeoutHandler
            End If
            
            ' ���N�G�X�g�Ԋu�������_����
            SafeSleep GetOptimizedWaitTime(800, 1500)
            
            If storeAddress = stAddr Then
                toStoreMin = 0
                storeOK = True
            Else
                Dim r3 As Variant: r3 = GetTransitInfo(driver, stAddr, storeAddress)
                If IsArray(r3) Then
                    toStoreMin = r3(0)
                    ParseStations r3(1), sDep, sArr
                    sText = r3(2)
                    If toStoreMin <= grpTime Then
                        storeOK = True
                    End If
                End If
            End If
            
            isValid = baseOK And storeOK
            
        ElseIf checkFlg = 3 Then
            ' �X�܂���̂ݔ���
            If storeAddress = stAddr Then
                toStoreMin = 0
                isValid = True
            Else
                Dim r4 As Variant: r4 = GetTransitInfo(driver, stAddr, storeAddress)
                If IsArray(r4) Then
                    toStoreMin = r4(0)
                    ParseStations r4(1), sDep, sArr
                    sText = r4(2)
                    If toStoreMin <= grpTime Then
                        isValid = True
                    End If
                End If
            End If
        End If
        
        processedCount = processedCount + 1
        
        ' �����𖞂����ŏ��̌����̗p
        If isValid Then
            foundIndex = idx
            depStation = Dep
            arrStation = Arr
            timeText = tText
            storeDepStation = sDep
            storeArrStation = sArr
            storeTimeText = sText
            Exit For
        End If
        
NextStation:
        ' ���������i��萔����������x�e�j
        If processedCount Mod 5 = 0 Then
            SafeSleep GetOptimizedWaitTime(2000, 3000) ' ���߂̋x�e
        End If
    Next i
    
    ' ���ʐݒ�
    If foundIndex >= 0 Then
        ' �ŏI�`�F�b�N - �������ݒ��O�ɂ�A�񂪋󂩂��m�F
        If Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value <> "" Then
            GetYahooTransit_UseOptionClick_WithDriver = False
            GoTo Finalize
        End If
        
        ' �V�[�g�ɏ�������
        Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value = dayToSet
        
        Select Case checkFlg
        Case 1
            SaveTransitInfo foundIndex, depStation, arrStation, timeText, 0
            GetYahooTransit_UseOptionClick_WithDriver = True
        Case 3
            SaveTransitInfo foundIndex, storeDepStation, storeArrStation, storeTimeText, 1
            GetYahooTransit_UseOptionClick_WithDriver = True
        Case 2
            SaveTransitInfo foundIndex, depStation, arrStation, timeText, 0
            SaveTransitInfo foundIndex, storeDepStation, storeArrStation, storeTimeText, 1
            GetYahooTransit_UseOptionClick_WithDriver = True
        End Select
        
    Else
        GetYahooTransit_UseOptionClick_WithDriver = False
    End If
    
    GoTo Finalize

TimeoutHandler:
    ' �^�C���A�E�g���̏���
    Call WriteLog("GetYahooTransit: �^�C���A�E�g���� - �����𒆒f���܂�", "GetYahooTransit_UseOptionClick_WithDriver")
    
    ' �u���E�U���Z�b�g - ������
    On Error Resume Next
    driver.ExecuteScript "window.stop();"
    driver.NavigateTo "about:blank"
    SafeSleep 2000 ' ���߂̑ҋ@
    
    ' �u���E�U��Ԃ̃N���[���A�b�v������
    driver.ExecuteScript "window.localStorage.clear(); window.sessionStorage.clear();"
    SafeSleep 1000
    On Error GoTo 0
    
    GetYahooTransit_UseOptionClick_WithDriver = False

Finalize:
    ' �h���C�o�[�͕��Ȃ� - �Ăяo�����ŊǗ�
End Function

' �^�C���A�E�g�񐔃J�E���g�p�֐��i�V�K�ǉ��j
Public Function IncrementTimeoutCounter(ByRef counter As Integer, ByVal maxTimeouts As Integer) As Boolean
    counter = counter + 1
    If counter >= maxTimeouts Then
        Call WriteLog("�ő�^�C���A�E�g��(" & maxTimeouts & ")�ɒB���܂����B���[�v�𔲂��܂��B", "TimeoutHandler")
        IncrementTimeoutCounter = True ' ���[�v�𔲂���
    Else
        Call WriteLog("�^�C���A�E�g��: " & counter & "/" & maxTimeouts, "TimeoutHandler")
        IncrementTimeoutCounter = False ' �p��
    End If
End Function













'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'�� ������ʋ@�ւł̌����ƍœK�ȃX�e�[�V�����̑I���i���S�ŁE�����h���C�o�[�g�p�j
'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
Public Function GetYahooTransit_UseOptionClick_Safe_WithDriver( _
    ByVal driver As SeleniumVBA.WebDriver, _
    ByVal baseAddress As String, _
    ByVal dayToSet As String, _
    ByVal checkFlg As Integer, _
    ByVal grpTime As Integer, _
    Optional ByVal storeAddress As String = "", _
    Optional ByRef usedLabels As Collection = Nothing, _
    Optional ByRef excludedStations As String = "" _
) As Boolean

    ' A��Ɋ��ɒl������Z���i�ǂ̃��x���ł��j�ւ̐ݒ��h�~����`�F�b�N
    Dim loopMAX As Long
    loopMAX = Application.WorksheetFunction.CountIf(Sheets(COMMON_SETTING7).Range("E4:E43"), "*")
    
    Dim r As Long
    ' �������x���̏d���`�F�b�N
    For r = 0 To loopMAX - 1
        ' ���S��v�݂̂��`�F�b�N
        If Sheets(COMMON_SETTING7).cells(4 + r, "A").Value = dayToSet Then
            ' ���ɓ����j�����ݒ肳��Ă���
            GetYahooTransit_UseOptionClick_Safe_WithDriver = False
            Exit Function
        End If
    Next r
    
    ' �R���N�V�����ł̏d���`�F�b�N
    If Not usedLabels Is Nothing Then
        ' �G���[��������邽�߂̃V���v���ȃ��[�v�`�F�b�N
        On Error Resume Next
        Dim i As Long
        For i = 1 To usedLabels.COUNT
            If usedLabels(i) = dayToSet Then
                On Error GoTo 0
                GetYahooTransit_UseOptionClick_Safe_WithDriver = False
                Exit Function
            End If
        Next i
        On Error GoTo 0
        
        ' �g�p�\��Ƃ��Ēǉ�
        usedLabels.Add dayToSet
    End If
    
    ' �����h���C�o�[���g�p����֐����Ăяo��
    Dim result As Boolean
    result = GetYahooTransit_UseOptionClick_WithDriver(driver, baseAddress, dayToSet, checkFlg, grpTime, storeAddress, excludedStations)
    
    ' �Ăяo�������s�����ꍇ�A�R���N�V��������폜
    If Not result And Not usedLabels Is Nothing Then
        On Error Resume Next
        ' Count�Ԗڂ��Ō�ɒǉ���������
        usedLabels.Remove usedLabels.COUNT
        On Error GoTo 0
    End If
    
    GetYahooTransit_UseOptionClick_Safe_WithDriver = result
End Function



'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'�� ������ʋ@�ւł̌����ƍœK�ȃX�e�[�V�����̑I��
'  checkFlg: ���胂�[�h (1=��Z���̂�, 2=��Z�����X��, 3=�X�܂̂�)
'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
Public Function GetYahooTransit_UseOptionClick( _
    ByVal baseAddress As String, _
    ByVal dayToSet As String, _
    ByVal checkFlg As Integer, _
    ByVal grpTime As Integer, _
    Optional ByVal storeAddress As String = "", _
    Optional ByVal excludedStations As String = "" _
) As Boolean

    Dim driver              As SeleniumVBA.WebDriver
    Dim keys                As SeleniumVBA.WebKeyboard
    Dim loopMAX             As Integer
    Dim foundIndex          As Integer  ' �����𖞂����ŏ��̃C���f�b�N�X
    Dim i                   As Integer
    Dim processedCount      As Integer
    
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    '���ۑ��p
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    Dim depStation          As String, arrStation As String, timeText As String
    Dim storeDepStation     As String, storeArrStation As String, storeTimeText As String
    
    On Error GoTo Finalize
    
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    'WebDriver������
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    driver.StartEdge
    driver.OpenBrowser
    
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    '�V�[�g2����X�e�[�V�������擾
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    With Sheets(COMMON_SETTING7)
        loopMAX = WorksheetFunction.CountIf(.Range("E4:E43"), "*")
    End With
    
    '������
    foundIndex = -1
    processedCount = 0
    depStation = "": arrStation = "": timeText = ""
    storeDepStation = "": storeArrStation = "": storeTimeText = ""
    
    ' ���O�X�e�[�V�����z��̍쐬
    Dim excludedArray() As String
    Dim excludedCount As Integer
    excludedCount = 0
    
    If Len(Trim(excludedStations)) > 0 Then
        excludedArray = Split(excludedStations, ",")
        excludedCount = UBound(excludedArray) + 1
    End If
    
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    '�ΏۃX�e�[�V�����z��쐬�i�œK���j- A�񂪋�̂��̂�����Ώ�
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    Dim targetStations() As Integer
    ReDim targetStations(loopMAX - 1)
    Dim targetCount As Integer: targetCount = 0
    
    For i = 0 To loopMAX - 1
        If Sheets(COMMON_SETTING7).Range("A" & 4 + i).Value = "" Then
            Dim stNum As String
            stNum = CStr(Sheets(COMMON_SETTING7).Range("E" & 4 + i).Value)
            
            ' ���O�X�e�[�V�����`�F�b�N
            Dim isExcluded As Boolean
            isExcluded = False
            
            If excludedCount > 0 Then
                Dim e As Integer
                For e = 0 To excludedCount - 1
                    If stNum = excludedArray(e) Then
                        isExcluded = True
                        Exit For
                    End If
                Next e
            End If
            
            If Not isExcluded Then
                targetStations(targetCount) = i
                targetCount = targetCount + 1
            End If
        End If
    Next i
    
    If targetCount = 0 Then
        GetYahooTransit_UseOptionClick = False
        GoTo Finalize
    End If
    
    ReDim Preserve targetStations(targetCount - 1)
    
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    '�ΏۃX�e�[�V���������������i30���ȓ��̍ŏ��̌����̗p�j
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    For i = 0 To targetCount - 1
        Dim idx As Integer: idx = targetStations(i)
        
        '���N�G�X�g�ԑҋ@
        If processedCount > 0 Then SafeSleep 500 + Int(Rnd() * 500) '500-1000ms   '�����_�����l
        
        Dim stAddr As String: stAddr = Sheets(COMMON_SETTING7).Range("G" & 4 + idx).Value
        If stAddr = "" Then GoTo NextStation
        
        ' �ēxA�񂪋󂩃`�F�b�N�i���̏����Őݒ肳��Ă��Ȃ����j
        If Sheets(COMMON_SETTING7).Range("A" & 4 + idx).Value <> "" Then GoTo NextStation
        
        Dim toBaseMin  As Integer: toBaseMin = 999
        Dim toStoreMin As Integer: toStoreMin = 999
        Dim Dep        As String, Arr As String, tText As String
        Dim sDep       As String, sArr As String, sText As String
        
        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
        '���򏈗� - �d�l�ʂ�30���ȓ��Ȃ瑦�̗p
        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
        Dim isValid As Boolean: isValid = False
        
        If checkFlg = 1 Then
            '��Z������̂ݔ���
            If baseAddress = stAddr Then
                toBaseMin = 0
                isValid = True
            Else
                Dim r1 As Variant: r1 = GetTransitInfo(driver, baseAddress, stAddr)
                If IsArray(r1) Then
                    toBaseMin = r1(0)
                    ParseStations r1(1), Dep, Arr
                    tText = r1(2)
                    If toBaseMin <= grpTime Then
                        isValid = True
                    End If
                End If
            End If
            
        ElseIf checkFlg = 2 Then
            '��Z�����X�e�[�V���� & �X�e�[�V�������X�� ��������
            Dim baseOK As Boolean: baseOK = False
            Dim storeOK As Boolean: storeOK = False
            
            If baseAddress = stAddr Then
                toBaseMin = 0
                baseOK = True
            Else
                Dim r2 As Variant: r2 = GetTransitInfo(driver, baseAddress, stAddr)
                If IsArray(r2) Then
                    toBaseMin = r2(0)
                    ParseStations r2(1), Dep, Arr
                    tText = r2(2)
                    If toBaseMin <= grpTime Then
                        baseOK = True
                    End If
                End If
            End If
            
            ' �x�[�X�����������Ă��Ȃ��Ȃ�X�L�b�v
            If Not baseOK Then GoTo NextStation
            
            SafeSleep 500 + Int(Rnd() * 500)  ' �Z�k
            
            If storeAddress = stAddr Then
                toStoreMin = 0
                storeOK = True
            Else
                Dim r3 As Variant: r3 = GetTransitInfo(driver, stAddr, storeAddress)
                If IsArray(r3) Then
                    toStoreMin = r3(0)
                    ParseStations r3(1), sDep, sArr
                    sText = r3(2)
                    If toStoreMin <= grpTime Then
                        storeOK = True
                    End If
                End If
            End If
            
            isValid = baseOK And storeOK
            
        ElseIf checkFlg = 3 Then
            '�X�܂���̂ݔ���
            If storeAddress = stAddr Then
                toStoreMin = 0
                isValid = True
            Else
                Dim r4 As Variant: r4 = GetTransitInfo(driver, stAddr, storeAddress)
                If IsArray(r4) Then
                    toStoreMin = r4(0)
                    ParseStations r4(1), sDep, sArr
                    sText = r4(2)
                    If toStoreMin <= grpTime Then
                        isValid = True
                    End If
                End If
            End If
        End If
        
        processedCount = processedCount + 1
        
        '�����𖞂����ŏ��̌����̗p�i�d�v�ȕύX�_�j
        If isValid Then
            foundIndex = idx
            depStation = Dep
            arrStation = Arr
            timeText = tText
            storeDepStation = sDep
            storeArrStation = sArr
            storeTimeText = sText
            Exit For  ' �����ɍ��v�����瑦�I��
        End If
        
NextStation:
    Next i
    
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    '���ʐݒ�
    '�\�\�\�\�\�\�\�\�\�\�\�\�\�\
    If foundIndex >= 0 Then
        ' �ŏI�`�F�b�N - �������ݒ��O�ɂ�A�񂪋󂩂��m�F
        If Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value <> "" Then
            GetYahooTransit_UseOptionClick = False
            GoTo Finalize
        End If
        
        ' �V�[�g�ɏ�������
        Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value = dayToSet
        
        Select Case checkFlg
        Case 1
            SaveTransitInfo foundIndex, depStation, arrStation, timeText, 0
            GetYahooTransit_UseOptionClick = True
        Case 3
            SaveTransitInfo foundIndex, storeDepStation, storeArrStation, storeTimeText, 1  ' FV/FW/FX��
            GetYahooTransit_UseOptionClick = True
        Case 2
            SaveTransitInfo foundIndex, depStation, arrStation, timeText, 0
            SaveTransitInfo foundIndex, storeDepStation, storeArrStation, storeTimeText, 1
            GetYahooTransit_UseOptionClick = True
        End Select
        
    Else
        GetYahooTransit_UseOptionClick = False
    End If

Finalize:
    If Not driver Is Nothing Then
        On Error Resume Next
        driver.CloseBrowser
        driver.Shutdown
    End If
End Function

'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'�� ������ʋ@�ւł̌����ƍœK�ȃX�e�[�V�����̑I���i���S�Łj
'  ����GetYahooTransit_UseOptionClick�����b�v���ďd���`�F�b�N�@�\��ǉ�
'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
Public Function GetYahooTransit_UseOptionClick_Safe( _
    ByVal baseAddress As String, _
    ByVal dayToSet As String, _
    ByVal checkFlg As Integer, _
    ByVal grpTime As Integer, _
    Optional ByVal storeAddress As String = "", _
    Optional ByRef usedLabels As Collection = Nothing, _
    Optional ByRef excludedStations As String = "" _
) As Boolean

    ' A��Ɋ��ɒl������Z���i�ǂ̃��x���ł��j�ւ̐ݒ��h�~����`�F�b�N
    Dim loopMAX As Long
    loopMAX = Application.WorksheetFunction.CountIf(Sheets(COMMON_SETTING7).Range("E4:E43"), "*")
    
    Dim r As Long
    ' �������x���̏d���`�F�b�N
    For r = 0 To loopMAX - 1
        ' ���S��v�݂̂��`�F�b�N
        If Sheets(COMMON_SETTING7).cells(4 + r, "A").Value = dayToSet Then
            ' ���ɓ����j�����ݒ肳��Ă���
            GetYahooTransit_UseOptionClick_Safe = False
            Exit Function
        End If
    Next r
    
    ' �R���N�V�����ł̏d���`�F�b�N
    If Not usedLabels Is Nothing Then
        ' �G���[��������邽�߂̃V���v���ȃ��[�v�`�F�b�N
        On Error Resume Next
        Dim i As Long
        For i = 1 To usedLabels.COUNT
            If usedLabels(i) = dayToSet Then
                On Error GoTo 0
                GetYahooTransit_UseOptionClick_Safe = False
                Exit Function
            End If
        Next i
        On Error GoTo 0
        
        ' �g�p�\��Ƃ��Ēǉ�
        usedLabels.Add dayToSet
    End If
    
    ' �����̊֐����Ăяo�� - ���O�X�e�[�V��������ǉ�
    Dim result As Boolean
    result = GetYahooTransit_UseOptionClick(baseAddress, dayToSet, checkFlg, grpTime, storeAddress, excludedStations)
    
    ' �Ăяo�������s�����ꍇ�A�R���N�V��������폜
    If Not result And Not usedLabels Is Nothing Then
        On Error Resume Next
        ' Count�Ԗڂ��Ō�ɒǉ���������
        usedLabels.Remove usedLabels.COUNT
        On Error GoTo 0
    End If
    
    GetYahooTransit_UseOptionClick_Safe = result
End Function

'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'�� �w���X�g�u�o�����o�R�������v����o���w�Ɠ����w���擾
'  stList: "A �� B �� C"
'  Dep, Arr �� ByRef �ŕԂ�
'  ���p�w������ꍇ�͏o���w�Ɋ��ʕt���ŒǋL
'�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
Private Sub ParseStations(ByVal stList As String, ByRef Dep As String, ByRef Arr As String)
    Dim arrSt As Variant
    Dep = "": Arr = ""
    If stList <> "" Then
        arrSt = Split(stList, " �� ")
        If UBound(arrSt) >= 0 Then
            ' �o���w�ݒ�
            Dep = arrSt(0)
            
            ' ���p�w������ꍇ�͊��ʕt���Œǉ�
            If UBound(arrSt) > 1 Then
                Dim transferStations As String
                Dim i As Integer
                
                ' �ŏ��̒��p�w(�C���f�b�N�X1)����J�n
                transferStations = arrSt(1)
                
                ' �c��̒��p�w��ǉ�
                For i = 2 To UBound(arrSt) - 1
                    transferStations = transferStations & "��" & arrSt(i)
                Next i
                
                ' �o���w�ɒ��p�w�������ʕt���Œǉ�
                Dep = Dep & "�i" & transferStations & "�j"
            End If
            
            ' �����w�ݒ�
            Arr = arrSt(UBound(arrSt))
        End If
    End If
End Sub


' �H�������V�[�g�ɕۑ�����֐�
' columnGroup: 0=FS/FT/FU��ɕۑ�, 1=FV/FW/FX��ɕۑ�
Public Sub SaveTransitInfo(ByVal stationIndex As Integer, ByVal departureStation As String, ByVal arrivalStation As String, ByVal travelTime As String, Optional ByVal columnGroup As Integer = 0)
    ' �C���f�b�N�X��4�������ăV�[�g�̍s�ԍ����擾
    Dim rowNum As Integer
    rowNum = stationIndex + 4
    
    ' �ۑ���̗������
    Dim departureCol As String, arrivalCol As String, timeCol As String
    
    If columnGroup = 0 Then
        ' �X�e�[�V�����Ԃ̏�� (FS/FT/FU)
        departureCol = "FS"
        arrivalCol = "FT"
        timeCol = "FU"
    Else
        ' �X��-�X�e�[�V�����Ԃ̏�� (FV/FW/FX)
        departureCol = "FV"
        arrivalCol = "FW"
        timeCol = "FX"
    End If
    
    ' Sheet2�ɏ����L�^
    With Sheets(COMMON_SETTING7)
        .cells(rowNum, departureCol).Value = departureStation  ' �o���w�i�抷�w���܂ށj
        .cells(rowNum, arrivalCol).Value = arrivalStation      ' �����w
        .cells(rowNum, timeCol).Value = travelTime            ' �ړ�����
    End With
End Sub


' ������ʋ@�ւł̈ړ����Ԃ��擾����֐�
' �����̊֐��͊����̏����Ƃ̌݊����̂��߂Ɉێ�
Private Function GetTransitTime(ByVal driver As SeleniumVBA.WebDriver, ByVal fromAddress As String, ByVal toAddress As String) As Integer
    ' �o���n�ƖړI�n�������ꍇ��0��Ԃ�
    If fromAddress = toAddress Then
        GetTransitTime = 0
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' Yahoo!�H�����ɃA�N�Z�X
    driver.NavigateTo "https://transit.yahoo.co.jp/"
    
    ' �y�[�W�ǂݍ��݊������m�F�i�Œ�ҋ@�ł͂Ȃ��v�f���\�������܂őҋ@�j
    Dim fromInput As SeleniumVBA.WebElement
    Set fromInput = WaitFindElement(driver, "input[name='from']", 5)
    If fromInput Is Nothing Then
        GetTransitTime = 999
        Exit Function
    End If
    
    ' �Z�������
    fromInput.Clear ' ���̓t�B�[���h���N���A
    fromInput.SendKeys fromAddress
    
    Dim toInput As SeleniumVBA.WebElement
    Set toInput = driver.FindElement(By.cssSelector, "input[name='to']")
    toInput.Clear ' ���̓t�B�[���h���N���A
    toInput.SendKeys toAddress
    
    ' �o��������9:00�ɐݒ�
    ' ���Ԃ�9���ɐݒ�i�h���b�v�_�E���N���b�N�̑O�ɏ����ҋ@�j
    SafeSleep 300
    Dim selHH As SeleniumVBA.WebElement
    Set selHH = driver.FindElement(By.cssSelector, "select[name='hh']")
    If Not selHH Is Nothing Then
        selHH.Click
        SafeSleep 300 ' �ҋ@���Ԃ�Z�k
        
        ' <option value="09">���N���b�N
        Dim optHH As SeleniumVBA.WebElement
        Set optHH = selHH.FindElement(By.cssSelector, "option[value='09']")
        If Not optHH Is Nothing Then
            optHH.Click
        End If
    End If
    
    ' ����00���ɐݒ�
    SafeSleep 300
    Dim selMM As SeleniumVBA.WebElement
    Set selMM = driver.FindElement(By.cssSelector, "select[name='mm']")
    If Not selMM Is Nothing Then
        selMM.Click
        SafeSleep 300 ' �ҋ@���Ԃ�Z�k
        
        ' <option value="00">���N���b�N
        Dim optMM As SeleniumVBA.WebElement
        Set optMM = selMM.FindElement(By.cssSelector, "option[value='00']")
        If Not optMM Is Nothing Then
            optMM.Click
        End If
    End If
    
    ' �����{�^�����N���b�N
    driver.FindElement(By.ID, "searchModuleSubmit").Click
    
    ' ���ʂ��\�������܂őҋ@�i�Œ�ҋ@�ł͂Ȃ����ʗv�f��ҋ@�j
    Dim summaryBlock As SeleniumVBA.WebElement
    Set summaryBlock = WaitFindElement(driver, "div#route01 div.routeSummary ul.summary", 8)
    If summaryBlock Is Nothing Then
        GetTransitTime = 999 ' ������Ȃ��ꍇ�͑傫�Ȓl��Ԃ�
        Exit Function
    End If
    
    ' ���v����(li.time)���擾
    Dim timeEl As SeleniumVBA.WebElement
    On Error Resume Next
    Set timeEl = summaryBlock.FindElement(By.cssSelector, "li.time")
    On Error GoTo 0
    
    If timeEl Is Nothing Then
        GetTransitTime = 999 ' ������Ȃ��ꍇ�͑傫�Ȓl��Ԃ�
        Exit Function
    End If
    
    Dim rawTimeText As String
    rawTimeText = timeEl.GetText
    Dim totalTimeText As String
    totalTimeText = ExtractTimeOnly(rawTimeText)
    
    ' ���ԃe�L�X�g�𕪒P�ʂɕϊ�
    GetTransitTime = ConvertTimeTextToMinutes(totalTimeText)
    Exit Function
    
ErrorHandler:
    GetTransitTime = 999 ' �G���[�����������ꍇ�͑傫�Ȓl��Ԃ�
End Function

' ���ԃe�L�X�g�i��F"1����05��"�j�𕪒P�ʂɕϊ�����֐�
Private Function ConvertTimeTextToMinutes(ByVal timeText As String) As Integer
    Dim hours As Integer
    Dim minutes As Integer
    
    ' ������
    hours = 0
    minutes = 0
    
    ' ���Ԃ��܂܂�Ă��邩�`�F�b�N
    If InStr(timeText, "����") > 0 Then
        ' ���ԕ����𒊏o
        Dim hoursPart As String
        hoursPart = Split(timeText, "����")(0)
        If IsNumeric(hoursPart) Then
            hours = CInt(hoursPart)
        End If
        
        ' �������݂���ꍇ�͒��o
        If InStr(timeText, "��") > 0 Then
            Dim minutesPart As String
            minutesPart = Split(Split(timeText, "����")(1), "��")(0)
            If IsNumeric(minutesPart) Then
                minutes = CInt(minutesPart)
            End If
        End If
    ' ���݂̂̏ꍇ
    ElseIf InStr(timeText, "��") > 0 Then
        Dim onlyMinutes As String
        onlyMinutes = Split(timeText, "��")(0)
        If IsNumeric(onlyMinutes) Then
            minutes = CInt(onlyMinutes)
        End If
    End If
    
    ConvertTimeTextToMinutes = hours * 60 + minutes
End Function

'-----------------------------------------------------------------------------
' ���v�f�ҋ@�֐�: �w��CSS�Z���N�^��������܂ŁA�ő�timeoutSec�b���g���C����
'-----------------------------------------------------------------------------
Private Function WaitFindElement(ByVal drv As SeleniumVBA.WebDriver, _
                                 ByVal cssSel As String, _
                                 ByVal timeoutSec As Long) As SeleniumVBA.WebElement
    Dim tEnd As Date
    tEnd = DateAdd("s", timeoutSec, Now)
    
    Dim el As SeleniumVBA.WebElement
    
    Do
        On Error Resume Next
        Set el = drv.FindElement(By.cssSelector, cssSel)
        On Error GoTo 0
        
        If Not el Is Nothing Then Exit Do
        DoEvents
    Loop While Now < tEnd
    
    If Not el Is Nothing Then
        Set WaitFindElement = el
    Else
        Set WaitFindElement = Nothing
    End If
End Function

'-----------------------------------------------------------------------------
' �����v���Ԃ̕�����𒊏o����֐�
'   ��: "09:00����10:05�� 1����05���i���40���j" �� "1����05��"
'-----------------------------------------------------------------------------
Private Function ExtractTimeOnly(ByVal rawText As String) As String
    Dim pos As Long
    Dim resultText As String
    Dim parenPos As Long
    
    ' �u���v�̈ʒu��T���i�X�y�[�X�̗L������Ȃ��j
    pos = InStr(rawText, "��")
    If pos > 0 Then
        resultText = Mid(rawText, pos + Len("��"))
    Else
        resultText = rawText
    End If
    
    ' �u�i�v������΂���ȍ~���폜
    parenPos = InStr(resultText, "�i")
    If parenPos > 0 Then
        resultText = Left(resultText, parenPos - 1)
    End If
    
    ExtractTimeOnly = Trim(resultText)
End Function

' �d�Ԉړ����Ԏ擾�֐�
Public Function GetTravelTimeMinutes(ByVal fromAddr As String, ByVal toAddr As String, Optional ByVal storeAddr As String = "") As Integer
    Dim driver As SeleniumVBA.WebDriver
    Dim result As Integer
    
    Set driver = SeleniumVBA.New_WebDriver
    driver.StartEdge
    driver.OpenBrowser
    
    If storeAddr = "" Then
        ' �X�e�[�V�����Ԉړ�
        result = GetTransitTime(driver, fromAddr, toAddr)
    Else
        ' �X�e�[�V�������X�܈ړ�
        Dim toStoreTime As Integer
        toStoreTime = GetTransitTime(driver, toAddr, storeAddr)
        If toStoreTime > 0 Then
            result = GetTransitTime(driver, fromAddr, toAddr) + toStoreTime
        Else
            result = 999 ' �G���[
        End If
    End If
    
    driver.CloseBrowser
    driver.Shutdown
    
    GetTravelTimeMinutes = result
End Function



