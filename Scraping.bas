Attribute VB_Name = "Module5"
' ScrapingModuleの先頭に追加
Public Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)

'--- スクレイピングモジュールの最適化 ---
' キャッシュの有効期限設定（セッション内でのみ有効）
Private CacheEnabled As Boolean
Private MaxCacheSize As Long
Private TransitCache As Object ' Dictionary型

' キャッシュの初期化 - 最適化版
Private Sub InitializeCache()
    If TransitCache Is Nothing Then
        Set TransitCache = CreateObject("Scripting.Dictionary")
        CacheEnabled = True
        MaxCacheSize = 6000  '
    End If
End Sub

' 結果をキャッシュ - 最適化版
Private Sub CacheTransitResult(fromAddr As String, toAddr As String, result As Variant)
    If Not CacheEnabled Then Exit Sub
    
    InitializeCache
    Dim key As String
    key = MakeCacheKey(fromAddr, toAddr)
    
    ' キャッシュサイズ制限
    If TransitCache.COUNT >= MaxCacheSize Then
        ' 最も古いエントリを削除（簡易実装）
        Dim firstKey As Variant
        firstKey = TransitCache.keys()(0)
        TransitCache.Remove firstKey
    End If
    
    If Not TransitCache.Exists(key) Then
        TransitCache.Add key, result
    End If
End Sub

' スクレイピング最適化待機時間取得
Public Function GetOptimizedWaitTime(minMs As Integer, maxMs As Integer) As Integer
    ' 通常より少し短い待機時間を返す
    GetOptimizedWaitTime = minMs + Int(Rnd() * (maxMs - minMs + 1))
End Function


' キャッシュキー作成（出発地+目的地）
Private Function MakeCacheKey(fromAddr As String, toAddr As String) As String
    MakeCacheKey = fromAddr & "→" & toAddr
End Function


' キャッシュ確認
Private Function GetCachedTransitResult(fromAddr As String, toAddr As String) As Variant
    InitializeCache
    Dim key As String
    key = MakeCacheKey(fromAddr, toAddr)
    
    ' 通常方向のキャッシュを確認
    If TransitCache.Exists(key) Then
        GetCachedTransitResult = TransitCache(key)
        Exit Function
    End If
    
    ' 逆方向のキャッシュを確認
    Dim reverseKey As String
    reverseKey = GetReverseCacheKey(fromAddr, toAddr)
    
    If TransitCache.Exists(reverseKey) Then
        Dim reverseResult As Variant
        reverseResult = TransitCache(reverseKey)
        
        ' 逆方向の結果を調整して返す
        Dim adjustedResult As Variant
        adjustedResult = AdjustReverseTransitResult(reverseResult)
        
        If Not IsNull(adjustedResult) Then
            ' 調整した結果を通常方向のキーでもキャッシュに保存
            If Not TransitCache.Exists(key) Then
                TransitCache.Add key, adjustedResult
            End If
            
            GetCachedTransitResult = adjustedResult
            Exit Function
        End If
    End If
    
    ' キャッシュに見つからない場合はNull
    GetCachedTransitResult = Null
End Function

' 逆方向のキーを作成する関数（新規作成）
Private Function GetReverseCacheKey(fromAddr As String, toAddr As String) As String
    GetReverseCacheKey = toAddr & "→" & fromAddr
End Function

' 逆方向の結果を調整する関数（新規作成）
Private Function AdjustReverseTransitResult(reverseResult As Variant) As Variant
    Dim adjustedResult(2) As Variant
    
    If IsArray(reverseResult) Then
        ' 時間はそのまま使用（往復でほぼ同じと仮定）
        adjustedResult(0) = reverseResult(0)
        
        ' 駅リストを逆順に変換
        Dim originalStations As String
        originalStations = reverseResult(1)
        
        If originalStations <> "" Then
            Dim stationArray As Variant
            stationArray = Split(originalStations, " → ")
            
            ' 配列を逆順にする
            Dim reversedStations As String
            reversedStations = ""
            
            Dim i As Integer
            For i = UBound(stationArray) To 0 Step -1
                If reversedStations = "" Then
                    reversedStations = stationArray(i)
                Else
                    reversedStations = reversedStations & " → " & stationArray(i)
                End If
            Next i
            
            adjustedResult(1) = reversedStations
        Else
            adjustedResult(1) = ""
        End If
        
        ' 時間テキストはそのまま使用
        adjustedResult(2) = reverseResult(2)
    Else
        ' 逆方向の結果が無効な場合はNullを返す
        AdjustReverseTransitResult = Null
        Exit Function
    End If
    
    AdjustReverseTransitResult = adjustedResult
End Function


' 改良版要素待機関数
Private Function WaitFindElementDynamic(ByVal drv As SeleniumVBA.WebDriver, _
                                       ByVal cssSel As String, _
                                       ByVal baseTimeout As Long) As SeleniumVBA.WebElement
    ' ランダムな追加時間
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
        ' 短い間隔でのポーリングを避けるために小さな待機を入れる
        drv.Wait 100
    Loop While Now < tEnd
    
    If Not el Is Nothing Then
        Set WaitFindElementDynamic = el
    Else
        Set WaitFindElementDynamic = Nothing
    End If
End Function


' 安全なスリープ関数（DoEvents呼び出しの削減）
Public Sub SafeSleep(ByVal ms As Long)
    ' DoEventsなしでWindowsAPI経由のスリープを実行
    SleepWinAPI ms
End Sub

' 安全な要素待機関数（タイムアウト機能強化）
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
        
        ' DoEvents呼び出しを制限（10回に1回）
        doEventsCounter = doEventsCounter + 1
        If doEventsCounter Mod 10 = 0 Then DoEvents
        
        ' より短い待機
        SleepWinAPI 100
    Loop While Now < EndTime
    
    Set SafeWaitFindElement = element
End Function

' 修正版 GetTransitInfo関数
Public Function GetTransitInfo(ByVal driver As SeleniumVBA.WebDriver, _
                              ByVal fromAddress As String, _
                              ByVal toAddress As String) As Variant
    Dim result(2) As Variant  ' 結果格納配列
    result(0) = 999           ' 所要時間（分）
    result(1) = ""            ' 駅リスト
    result(2) = ""            ' 所要時間テキスト
    
    ' 出発地と目的地が同じ場合は0を返す
    If fromAddress = toAddress Then
        result(0) = 0
        GetTransitInfo = result
        Exit Function
    End If
    
    ' キャッシュチェック
    Dim cachedResult As Variant
    cachedResult = GetCachedTransitResult(fromAddress, toAddress)
    If Not IsNull(cachedResult) Then
        GetTransitInfo = cachedResult
        Exit Function
    End If
    
    ' 最大実行時間の設定
    Dim maxExecutionTime As Long: maxExecutionTime = 30 ' 秒
    Dim StartTime As Date: StartTime = Now
    
    On Error GoTo ErrorHandler
    
    ' Yahoo路線情報にアクセス
    driver.NavigateTo "https://transit.yahoo.co.jp/"
    SafeSleep GetOptimizedWaitTime(600, 1000)
    
    ' 実行時間超過チェック
    If DateDiff("s", StartTime, Now) > maxExecutionTime Then GoTo TimeoutHandler
    
    ' 検索フォーム入力
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
    
    ' 実行時間超過チェック
    If DateDiff("s", StartTime, Now) > maxExecutionTime Then GoTo TimeoutHandler
    
    ' 時刻設定（9:00）
    Dim selHH As SeleniumVBA.WebElement
    Set selHH = driver.FindElement(By.cssSelector, "select[name='hh']")
    If Not selHH Is Nothing Then
        selHH.Click
        SafeSleep 300
        
        Dim optHH As SeleniumVBA.WebElement
        Set optHH = selHH.FindElement(By.cssSelector, "option[value='09']")
        If Not optHH Is Nothing Then optHH.Click
    End If
    
    ' 分設定（00分）
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
    
    ' 検索ボタンクリック
    SafeSleep 300
    driver.FindElement(By.ID, "searchModuleSubmit").Click
    
    ' 実行時間超過チェック
    If DateDiff("s", StartTime, Now) > maxExecutionTime Then GoTo TimeoutHandler
    
    ' 結果取得
    Dim summaryBlock As SeleniumVBA.WebElement
    Set summaryBlock = SafeWaitFindElement(driver, "div#route01 div.routeSummary ul.summary", 7)
    If summaryBlock Is Nothing Then GoTo ErrorHandler
    
    Dim timeEl As SeleniumVBA.WebElement
    On Error Resume Next
    Set timeEl = summaryBlock.FindElement(By.cssSelector, "li.time")
    On Error GoTo 0
    
    If timeEl Is Nothing Then GoTo ErrorHandler
    
    ' 時間テキスト取得・変換
    Dim rawTimeText As String: rawTimeText = timeEl.GetText
    Dim totalTimeText As String: totalTimeText = ExtractTimeOnly(rawTimeText)
    result(0) = ConvertTimeTextToMinutes(totalTimeText)
    result(2) = totalTimeText
    
    ' 駅名一覧取得
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
                stationList = stationList & " → " & stName
            End If
        Next
    End If
    
    result(1) = stationList
    CacheTransitResult fromAddress, toAddress, result
    
    GetTransitInfo = result
    Exit Function
    
TimeoutHandler:
    Call WriteLog("GetTransitInfo: タイムアウト - " & fromAddress & "→" & toAddress, "GetTransitInfo")
    On Error Resume Next
    driver.ExecuteScript "window.stop();"
    driver.NavigateTo "about:blank"
    SafeSleep 1000
    On Error GoTo 0
    GetTransitInfo = result
    Exit Function
    
ErrorHandler:
    Call WriteLog("GetTransitInfo: エラー - " & fromAddress & "→" & toAddress, "GetTransitInfo")
    GetTransitInfo = result
End Function

' GetYahooTransit_UseOptionClick_WithDriver関数
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
    
    ' 候補保存用変数
    Dim depStation          As String, arrStation As String, timeText As String
    Dim storeDepStation     As String, storeArrStation As String, storeTimeText As String
    
    On Error GoTo Finalize
    
    ' 最大実行時間設定
    maxExecutionTime = 120 ' 2分
    StartTime = Now
    
    ' シート2からステーション数取得
    With Sheets(COMMON_SETTING7)
        loopMAX = WorksheetFunction.CountIf(.Range("E4:E43"), "*")
    End With
    
    ' 初期化
    foundIndex = -1
    processedCount = 0
    depStation = "": arrStation = "": timeText = ""
    storeDepStation = "": storeArrStation = "": storeTimeText = ""
    
    ' 除外ステーション配列の作成
    Dim excludedArray() As String
    Dim excludedCount As Integer
    excludedCount = 0
    
    If Len(Trim(excludedStations)) > 0 Then
        excludedArray = Split(excludedStations, ",")
        excludedCount = UBound(excludedArray) + 1
    End If
    
    ' A列が空のステーションのみを対象にする
    Dim targetStations() As Integer
    ReDim targetStations(loopMAX - 1)
    Dim targetCount As Integer: targetCount = 0
    
    For i = 0 To loopMAX - 1
        ' タイムアウトチェック
        If DateDiff("s", StartTime, Now) > maxExecutionTime Then
            GoTo TimeoutHandler
        End If
        
        If Sheets(COMMON_SETTING7).Range("A" & 4 + i).Value = "" Then
            Dim stNum As String
            stNum = CStr(Sheets(COMMON_SETTING7).Range("E" & 4 + i).Value)
            
            ' 除外ステーションチェック
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
    
    ' 各ステーションを検索
    For i = 0 To targetCount - 1
        ' タイムアウトチェック
        If DateDiff("s", StartTime, Now) > maxExecutionTime Then
            GoTo TimeoutHandler
        End If
        
        Dim idx As Integer: idx = targetStations(i)
        
        ' リクエスト間待機
        If processedCount > 0 Then SafeSleep GetOptimizedWaitTime(1000, 2000)
        
        Dim stAddr As String: stAddr = Sheets(COMMON_SETTING7).Range("G" & 4 + idx).Value
        If stAddr = "" Then GoTo NextStation
        
        ' A列が空かを再確認
        If Sheets(COMMON_SETTING7).Range("A" & 4 + idx).Value <> "" Then GoTo NextStation
        
        Dim toBaseMin  As Integer: toBaseMin = 999
        Dim toStoreMin As Integer: toStoreMin = 999
        Dim Dep        As String, Arr As String, tText As String
        Dim sDep       As String, sArr As String, sText As String
        
        ' 分岐処理
        Dim isValid As Boolean: isValid = False
        
        If checkFlg = 1 Then
            ' 基準住所からのみ判定
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
            ' 基準住所→ステーション & ステーション→店舗 両方判定
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
            
            ' ベース条件満たしていないならスキップ
            If Not baseOK Then GoTo NextStation
            
            ' タイムアウトチェック
            If DateDiff("s", StartTime, Now) > maxExecutionTime Then
                GoTo TimeoutHandler
            End If
            
            ' リクエスト間隔をランダム化
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
            ' 店舗からのみ判定
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
        
        ' 条件を満たす最初の候補を採用
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
        ' 処理制限（一定数処理したら休憩）
        If processedCount Mod 5 = 0 Then
            SafeSleep GetOptimizedWaitTime(2000, 3000) ' 長めの休憩
        End If
    Next i
    
    ' 結果設定
    If foundIndex >= 0 Then
        ' 最終チェック - 書き込み直前にもA列が空かを確認
        If Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value <> "" Then
            GetYahooTransit_UseOptionClick_WithDriver = False
            GoTo Finalize
        End If
        
        ' シートに書き込み
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
    ' タイムアウト時の処理
    Call WriteLog("GetYahooTransit: タイムアウト発生 - 処理を中断します", "GetYahooTransit_UseOptionClick_WithDriver")
    
    ' ブラウザリセット - 強化版
    On Error Resume Next
    driver.ExecuteScript "window.stop();"
    driver.NavigateTo "about:blank"
    SafeSleep 2000 ' 長めの待機
    
    ' ブラウザ状態のクリーンアップを強化
    driver.ExecuteScript "window.localStorage.clear(); window.sessionStorage.clear();"
    SafeSleep 1000
    On Error GoTo 0
    
    GetYahooTransit_UseOptionClick_WithDriver = False

Finalize:
    ' ドライバーは閉じない - 呼び出し元で管理
End Function

' タイムアウト回数カウント用関数（新規追加）
Public Function IncrementTimeoutCounter(ByRef counter As Integer, ByVal maxTimeouts As Integer) As Boolean
    counter = counter + 1
    If counter >= maxTimeouts Then
        Call WriteLog("最大タイムアウト回数(" & maxTimeouts & ")に達しました。ループを抜けます。", "TimeoutHandler")
        IncrementTimeoutCounter = True ' ループを抜ける
    Else
        Call WriteLog("タイムアウト回数: " & counter & "/" & maxTimeouts, "TimeoutHandler")
        IncrementTimeoutCounter = False ' 継続
    End If
End Function













'―――――――――――――――――――――――――――――――
'■ 公共交通機関での検索と最適なステーションの選択（安全版・既存ドライバー使用）
'―――――――――――――――――――――――――――――――
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

    ' A列に既に値があるセル（どのラベルでも）への設定を防止するチェック
    Dim loopMAX As Long
    loopMAX = Application.WorksheetFunction.CountIf(Sheets(COMMON_SETTING7).Range("E4:E43"), "*")
    
    Dim r As Long
    ' 同じラベルの重複チェック
    For r = 0 To loopMAX - 1
        ' 完全一致のみをチェック
        If Sheets(COMMON_SETTING7).cells(4 + r, "A").Value = dayToSet Then
            ' 既に同じ曜日が設定されている
            GetYahooTransit_UseOptionClick_Safe_WithDriver = False
            Exit Function
        End If
    Next r
    
    ' コレクションでの重複チェック
    If Not usedLabels Is Nothing Then
        ' エラーを回避するためのシンプルなループチェック
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
        
        ' 使用予定として追加
        usedLabels.Add dayToSet
    End If
    
    ' 既存ドライバーを使用する関数を呼び出す
    Dim result As Boolean
    result = GetYahooTransit_UseOptionClick_WithDriver(driver, baseAddress, dayToSet, checkFlg, grpTime, storeAddress, excludedStations)
    
    ' 呼び出しが失敗した場合、コレクションから削除
    If Not result And Not usedLabels Is Nothing Then
        On Error Resume Next
        ' Count番目が最後に追加した項目
        usedLabels.Remove usedLabels.COUNT
        On Error GoTo 0
    End If
    
    GetYahooTransit_UseOptionClick_Safe_WithDriver = result
End Function



'―――――――――――――――――――――――――――――――
'■ 公共交通機関での検索と最適なステーションの選択
'  checkFlg: 判定モード (1=基準住所のみ, 2=基準住所＆店舗, 3=店舗のみ)
'―――――――――――――――――――――――――――――――
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
    Dim foundIndex          As Integer  ' 条件を満たす最初のインデックス
    Dim i                   As Integer
    Dim processedCount      As Integer
    
    '――――――――――――――
    '候補保存用
    '――――――――――――――
    Dim depStation          As String, arrStation As String, timeText As String
    Dim storeDepStation     As String, storeArrStation As String, storeTimeText As String
    
    On Error GoTo Finalize
    
    '――――――――――――――
    'WebDriver初期化
    '――――――――――――――
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    driver.StartEdge
    driver.OpenBrowser
    
    '――――――――――――――
    'シート2からステーション数取得
    '――――――――――――――
    With Sheets(COMMON_SETTING7)
        loopMAX = WorksheetFunction.CountIf(.Range("E4:E43"), "*")
    End With
    
    '初期化
    foundIndex = -1
    processedCount = 0
    depStation = "": arrStation = "": timeText = ""
    storeDepStation = "": storeArrStation = "": storeTimeText = ""
    
    ' 除外ステーション配列の作成
    Dim excludedArray() As String
    Dim excludedCount As Integer
    excludedCount = 0
    
    If Len(Trim(excludedStations)) > 0 Then
        excludedArray = Split(excludedStations, ",")
        excludedCount = UBound(excludedArray) + 1
    End If
    
    '――――――――――――――
    '対象ステーション配列作成（最適化）- A列が空のものだけを対象
    '――――――――――――――
    Dim targetStations() As Integer
    ReDim targetStations(loopMAX - 1)
    Dim targetCount As Integer: targetCount = 0
    
    For i = 0 To loopMAX - 1
        If Sheets(COMMON_SETTING7).Range("A" & 4 + i).Value = "" Then
            Dim stNum As String
            stNum = CStr(Sheets(COMMON_SETTING7).Range("E" & 4 + i).Value)
            
            ' 除外ステーションチェック
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
    
    '――――――――――――――
    '対象ステーションを順次検索（30分以内の最初の候補を採用）
    '――――――――――――――
    For i = 0 To targetCount - 1
        Dim idx As Integer: idx = targetStations(i)
        
        'リクエスト間待機
        If processedCount > 0 Then SafeSleep 500 + Int(Rnd() * 500) '500-1000ms   'ランダム数値
        
        Dim stAddr As String: stAddr = Sheets(COMMON_SETTING7).Range("G" & 4 + idx).Value
        If stAddr = "" Then GoTo NextStation
        
        ' 再度A列が空かチェック（他の処理で設定されていないか）
        If Sheets(COMMON_SETTING7).Range("A" & 4 + idx).Value <> "" Then GoTo NextStation
        
        Dim toBaseMin  As Integer: toBaseMin = 999
        Dim toStoreMin As Integer: toStoreMin = 999
        Dim Dep        As String, Arr As String, tText As String
        Dim sDep       As String, sArr As String, sText As String
        
        '――――――――――――――
        '分岐処理 - 仕様通り30分以内なら即採用
        '――――――――――――――
        Dim isValid As Boolean: isValid = False
        
        If checkFlg = 1 Then
            '基準住所からのみ判定
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
            '基準住所→ステーション & ステーション→店舗 両方判定
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
            
            ' ベース条件満たしていないならスキップ
            If Not baseOK Then GoTo NextStation
            
            SafeSleep 500 + Int(Rnd() * 500)  ' 短縮
            
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
            '店舗からのみ判定
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
        
        '条件を満たす最初の候補を採用（重要な変更点）
        If isValid Then
            foundIndex = idx
            depStation = Dep
            arrStation = Arr
            timeText = tText
            storeDepStation = sDep
            storeArrStation = sArr
            storeTimeText = sText
            Exit For  ' 条件に合致したら即終了
        End If
        
NextStation:
    Next i
    
    '――――――――――――――
    '結果設定
    '――――――――――――――
    If foundIndex >= 0 Then
        ' 最終チェック - 書き込み直前にもA列が空かを確認
        If Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value <> "" Then
            GetYahooTransit_UseOptionClick = False
            GoTo Finalize
        End If
        
        ' シートに書き込み
        Sheets(COMMON_SETTING7).Range("A" & 4 + foundIndex).Value = dayToSet
        
        Select Case checkFlg
        Case 1
            SaveTransitInfo foundIndex, depStation, arrStation, timeText, 0
            GetYahooTransit_UseOptionClick = True
        Case 3
            SaveTransitInfo foundIndex, storeDepStation, storeArrStation, storeTimeText, 1  ' FV/FW/FX列
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

'―――――――――――――――――――――――――――――――
'■ 公共交通機関での検索と最適なステーションの選択（安全版）
'  元のGetYahooTransit_UseOptionClickをラップして重複チェック機能を追加
'―――――――――――――――――――――――――――――――
Public Function GetYahooTransit_UseOptionClick_Safe( _
    ByVal baseAddress As String, _
    ByVal dayToSet As String, _
    ByVal checkFlg As Integer, _
    ByVal grpTime As Integer, _
    Optional ByVal storeAddress As String = "", _
    Optional ByRef usedLabels As Collection = Nothing, _
    Optional ByRef excludedStations As String = "" _
) As Boolean

    ' A列に既に値があるセル（どのラベルでも）への設定を防止するチェック
    Dim loopMAX As Long
    loopMAX = Application.WorksheetFunction.CountIf(Sheets(COMMON_SETTING7).Range("E4:E43"), "*")
    
    Dim r As Long
    ' 同じラベルの重複チェック
    For r = 0 To loopMAX - 1
        ' 完全一致のみをチェック
        If Sheets(COMMON_SETTING7).cells(4 + r, "A").Value = dayToSet Then
            ' 既に同じ曜日が設定されている
            GetYahooTransit_UseOptionClick_Safe = False
            Exit Function
        End If
    Next r
    
    ' コレクションでの重複チェック
    If Not usedLabels Is Nothing Then
        ' エラーを回避するためのシンプルなループチェック
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
        
        ' 使用予定として追加
        usedLabels.Add dayToSet
    End If
    
    ' 既存の関数を呼び出す - 除外ステーション情報を追加
    Dim result As Boolean
    result = GetYahooTransit_UseOptionClick(baseAddress, dayToSet, checkFlg, grpTime, storeAddress, excludedStations)
    
    ' 呼び出しが失敗した場合、コレクションから削除
    If Not result And Not usedLabels Is Nothing Then
        On Error Resume Next
        ' Count番目が最後に追加した項目
        usedLabels.Remove usedLabels.COUNT
        On Error GoTo 0
    End If
    
    GetYahooTransit_UseOptionClick_Safe = result
End Function

'―――――――――――――――――――――――――――――――
'■ 駅リスト「出発→経由→到着」から出発駅と到着駅を取得
'  stList: "A → B → C"
'  Dep, Arr は ByRef で返す
'  中継駅がある場合は出発駅に括弧付きで追記
'―――――――――――――――――――――――――――――――
Private Sub ParseStations(ByVal stList As String, ByRef Dep As String, ByRef Arr As String)
    Dim arrSt As Variant
    Dep = "": Arr = ""
    If stList <> "" Then
        arrSt = Split(stList, " → ")
        If UBound(arrSt) >= 0 Then
            ' 出発駅設定
            Dep = arrSt(0)
            
            ' 中継駅がある場合は括弧付きで追加
            If UBound(arrSt) > 1 Then
                Dim transferStations As String
                Dim i As Integer
                
                ' 最初の中継駅(インデックス1)から開始
                transferStations = arrSt(1)
                
                ' 残りの中継駅を追加
                For i = 2 To UBound(arrSt) - 1
                    transferStations = transferStations & "→" & arrSt(i)
                Next i
                
                ' 出発駅に中継駅情報を括弧付きで追加
                Dep = Dep & "（" & transferStations & "）"
            End If
            
            ' 到着駅設定
            Arr = arrSt(UBound(arrSt))
        End If
    End If
End Sub


' 路線情報をシートに保存する関数
' columnGroup: 0=FS/FT/FU列に保存, 1=FV/FW/FX列に保存
Public Sub SaveTransitInfo(ByVal stationIndex As Integer, ByVal departureStation As String, ByVal arrivalStation As String, ByVal travelTime As String, Optional ByVal columnGroup As Integer = 0)
    ' インデックスに4を加えてシートの行番号を取得
    Dim rowNum As Integer
    rowNum = stationIndex + 4
    
    ' 保存先の列を決定
    Dim departureCol As String, arrivalCol As String, timeCol As String
    
    If columnGroup = 0 Then
        ' ステーション間の情報 (FS/FT/FU)
        departureCol = "FS"
        arrivalCol = "FT"
        timeCol = "FU"
    Else
        ' 店舗-ステーション間の情報 (FV/FW/FX)
        departureCol = "FV"
        arrivalCol = "FW"
        timeCol = "FX"
    End If
    
    ' Sheet2に情報を記録
    With Sheets(COMMON_SETTING7)
        .cells(rowNum, departureCol).Value = departureStation  ' 出発駅（乗換駅を含む）
        .cells(rowNum, arrivalCol).Value = arrivalStation      ' 到着駅
        .cells(rowNum, timeCol).Value = travelTime            ' 移動時間
    End With
End Sub


' 公共交通機関での移動時間を取得する関数
' ※この関数は既存の処理との互換性のために維持
Private Function GetTransitTime(ByVal driver As SeleniumVBA.WebDriver, ByVal fromAddress As String, ByVal toAddress As String) As Integer
    ' 出発地と目的地が同じ場合は0を返す
    If fromAddress = toAddress Then
        GetTransitTime = 0
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' Yahoo!路線情報にアクセス
    driver.NavigateTo "https://transit.yahoo.co.jp/"
    
    ' ページ読み込み完了を確認（固定待機ではなく要素が表示されるまで待機）
    Dim fromInput As SeleniumVBA.WebElement
    Set fromInput = WaitFindElement(driver, "input[name='from']", 5)
    If fromInput Is Nothing Then
        GetTransitTime = 999
        Exit Function
    End If
    
    ' 住所を入力
    fromInput.Clear ' 入力フィールドをクリア
    fromInput.SendKeys fromAddress
    
    Dim toInput As SeleniumVBA.WebElement
    Set toInput = driver.FindElement(By.cssSelector, "input[name='to']")
    toInput.Clear ' 入力フィールドをクリア
    toInput.SendKeys toAddress
    
    ' 出発時刻を9:00に設定
    ' 時間を9時に設定（ドロップダウンクリックの前に少し待機）
    SafeSleep 300
    Dim selHH As SeleniumVBA.WebElement
    Set selHH = driver.FindElement(By.cssSelector, "select[name='hh']")
    If Not selHH Is Nothing Then
        selHH.Click
        SafeSleep 300 ' 待機時間を短縮
        
        ' <option value="09">をクリック
        Dim optHH As SeleniumVBA.WebElement
        Set optHH = selHH.FindElement(By.cssSelector, "option[value='09']")
        If Not optHH Is Nothing Then
            optHH.Click
        End If
    End If
    
    ' 分を00分に設定
    SafeSleep 300
    Dim selMM As SeleniumVBA.WebElement
    Set selMM = driver.FindElement(By.cssSelector, "select[name='mm']")
    If Not selMM Is Nothing Then
        selMM.Click
        SafeSleep 300 ' 待機時間を短縮
        
        ' <option value="00">をクリック
        Dim optMM As SeleniumVBA.WebElement
        Set optMM = selMM.FindElement(By.cssSelector, "option[value='00']")
        If Not optMM Is Nothing Then
            optMM.Click
        End If
    End If
    
    ' 検索ボタンをクリック
    driver.FindElement(By.ID, "searchModuleSubmit").Click
    
    ' 結果が表示されるまで待機（固定待機ではなく結果要素を待機）
    Dim summaryBlock As SeleniumVBA.WebElement
    Set summaryBlock = WaitFindElement(driver, "div#route01 div.routeSummary ul.summary", 8)
    If summaryBlock Is Nothing Then
        GetTransitTime = 999 ' 見つからない場合は大きな値を返す
        Exit Function
    End If
    
    ' 所要時間(li.time)を取得
    Dim timeEl As SeleniumVBA.WebElement
    On Error Resume Next
    Set timeEl = summaryBlock.FindElement(By.cssSelector, "li.time")
    On Error GoTo 0
    
    If timeEl Is Nothing Then
        GetTransitTime = 999 ' 見つからない場合は大きな値を返す
        Exit Function
    End If
    
    Dim rawTimeText As String
    rawTimeText = timeEl.GetText
    Dim totalTimeText As String
    totalTimeText = ExtractTimeOnly(rawTimeText)
    
    ' 時間テキストを分単位に変換
    GetTransitTime = ConvertTimeTextToMinutes(totalTimeText)
    Exit Function
    
ErrorHandler:
    GetTransitTime = 999 ' エラーが発生した場合は大きな値を返す
End Function

' 時間テキスト（例："1時間05分"）を分単位に変換する関数
Private Function ConvertTimeTextToMinutes(ByVal timeText As String) As Integer
    Dim hours As Integer
    Dim minutes As Integer
    
    ' 初期化
    hours = 0
    minutes = 0
    
    ' 時間が含まれているかチェック
    If InStr(timeText, "時間") > 0 Then
        ' 時間部分を抽出
        Dim hoursPart As String
        hoursPart = Split(timeText, "時間")(0)
        If IsNumeric(hoursPart) Then
            hours = CInt(hoursPart)
        End If
        
        ' 分が存在する場合は抽出
        If InStr(timeText, "分") > 0 Then
            Dim minutesPart As String
            minutesPart = Split(Split(timeText, "時間")(1), "分")(0)
            If IsNumeric(minutesPart) Then
                minutes = CInt(minutesPart)
            End If
        End If
    ' 分のみの場合
    ElseIf InStr(timeText, "分") > 0 Then
        Dim onlyMinutes As String
        onlyMinutes = Split(timeText, "分")(0)
        If IsNumeric(onlyMinutes) Then
            minutes = CInt(onlyMinutes)
        End If
    End If
    
    ConvertTimeTextToMinutes = hours * 60 + minutes
End Function

'-----------------------------------------------------------------------------
' ■要素待機関数: 指定CSSセレクタが見つかるまで、最大timeoutSec秒リトライする
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
' ■所要時間の文字列を抽出する関数
'   例: "09:00発→10:05着 1時間05分（乗車40分）" → "1時間05分"
'-----------------------------------------------------------------------------
Private Function ExtractTimeOnly(ByVal rawText As String) As String
    Dim pos As Long
    Dim resultText As String
    Dim parenPos As Long
    
    ' 「着」の位置を探す（スペースの有無を問わない）
    pos = InStr(rawText, "着")
    If pos > 0 Then
        resultText = Mid(rawText, pos + Len("着"))
    Else
        resultText = rawText
    End If
    
    ' 「（」があればそれ以降を削除
    parenPos = InStr(resultText, "（")
    If parenPos > 0 Then
        resultText = Left(resultText, parenPos - 1)
    End If
    
    ExtractTimeOnly = Trim(resultText)
End Function

' 電車移動時間取得関数
Public Function GetTravelTimeMinutes(ByVal fromAddr As String, ByVal toAddr As String, Optional ByVal storeAddr As String = "") As Integer
    Dim driver As SeleniumVBA.WebDriver
    Dim result As Integer
    
    Set driver = SeleniumVBA.New_WebDriver
    driver.StartEdge
    driver.OpenBrowser
    
    If storeAddr = "" Then
        ' ステーション間移動
        result = GetTransitTime(driver, fromAddr, toAddr)
    Else
        ' ステーション→店舗移動
        Dim toStoreTime As Integer
        toStoreTime = GetTransitTime(driver, toAddr, storeAddr)
        If toStoreTime > 0 Then
            result = GetTransitTime(driver, fromAddr, toAddr) + toStoreTime
        Else
            result = 999 ' エラー
        End If
    End If
    
    driver.CloseBrowser
    driver.Shutdown
    
    GetTravelTimeMinutes = result
End Function



