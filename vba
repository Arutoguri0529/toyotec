Function YoyakuPTN_THU_TRAIN(ByVal PTN As Integer, ByVal Stallday As String) As Integer
    Dim loopMAX      As Long
    Dim grpTime      As Long
    Dim RET          As Integer
    Dim stationNo(0 To 3) As String
    Dim baseIdx(0 To 3) As Integer
    Dim foundAll     As Boolean
    Dim labels       As Variant
    Dim r As Long, rr As Long, j As Long, k As Long, m As Long
    Dim candCar      As String
    Dim baseAddress  As String, storeAddress As String
    Dim dayToSet     As String
    Dim usedLabels   As Collection
    Dim foundFlag    As Boolean
    Dim excludedStations As String
    Dim excludedStage2 As String   ' ステージ2専用除外リスト
    Dim excludedStage3 As String   ' ステージ3専用除外リスト
    Dim usedAsFirst() As String  ' 1番目に設定済みのステーション記録
    Dim usedAsFirstCount As Integer
    Dim walkSearchCompleted As Boolean  ' 徒歩条件での検索完了フラグ
    Dim i As Integer  ' 外部ループのインデックス
    
    ' タイムアウト監視用変数
    Dim maxExecutionTime   As Long
    Dim startTime          As Date
    Dim timeoutCounter     As Integer
    Dim MAX_TIMEOUTS       As Integer
    
    maxExecutionTime = 300 ' 5分
    startTime = Now
    timeoutCounter = 0
    MAX_TIMEOUTS = 3

    Call WriteLog("木曜パターン設定 開始: PTN=" & PTN & " Stallday=" & Stallday, "YoyakuPTN_THU_TRAIN", "処理開始")

    ' 単一のWebDriverインスタンスを作成
    Dim driver As SeleniumVBA.WebDriver
    Set driver = SeleniumVBA.New_WebDriver

    On Error GoTo CleanUpDriver

    driver.StartEdge
    driver.OpenBrowser
    Call WriteLog("WebDriver 初期化完了", "YoyakuPTN_THU_TRAIN", "ドライバー初期化")

    '--- ラベル配列定義 ---
    Select Case PTN
        Case 2: labels = Array("月曜日②（電車）", "火曜日①（電車）", "火曜日④（電車）", "水曜日③（電車）")
        Case 3: labels = Array("火曜日②（電車）", "水曜日①（電車）", "水曜日④（電車）", "木曜日③（電車）")
        Case 4: labels = Array("水曜日②（電車）", "木曜日①（電車）", "木曜日④（電車）", "金曜日③（電車）")
        Case Else:
            Call WriteLog("不正なPTN値: " & PTN, "YoyakuPTN_THU_TRAIN", "パラメータエラー")
            GoTo CleanUpDriver
    End Select
    
    Call WriteLog("ラベル定義: " & labels(0) & ", " & labels(1) & ", " & labels(2) & ", " & labels(3), "YoyakuPTN_THU_TRAIN", "ラベル定義")
    storeAddress = Sheets("Sheet2").Range("G2").Value
    Call WriteLog("店舗住所: " & storeAddress, "YoyakuPTN_THU_TRAIN", "住所取得")
    Set usedLabels = New Collection
    excludedStations = ""
    excludedStage2 = ""    ' ステージ2専用除外リスト初期化
    excludedStage3 = ""    ' ステージ3専用除外リスト初期化

    Thu_ClearYoyaku PTN, Stallday
    Call WriteLog("既存ラベルをクリア", "YoyakuPTN_THU_TRAIN", "ラベルクリア")

    loopMAX = Application.WorksheetFunction.CountIf(Sheets("Sheet2").Range("E4:E43"), "*")
    Call WriteLog("ステーション数: " & loopMAX, "YoyakuPTN_THU_TRAIN", "データ取得")
    If loopMAX < 4 Then
        Call WriteLog("ステーション数不足: 最低4つ必要", "YoyakuPTN_THU_TRAIN", "データ不足")
        GoTo CleanUpDriver
    End If
    grpTime = CLng(Sheets("Sheet1").Range("B25").Value)
    Call WriteLog("徒歩閾値: " & grpTime & "分", "YoyakuPTN_THU_TRAIN", "閾値取得")
    foundAll = False: RET = 0
    
    ' 初期化
    ReDim usedAsFirst(loopMAX - 1)
    usedAsFirstCount = 0
    walkSearchCompleted = False
    Call WriteLog("初期化完了: 徒歩検索モード", "YoyakuPTN_THU_TRAIN", "初期化完了")
    
    ' タイムアウトチェック
    If DateDiff("s", startTime, Now) > maxExecutionTime Then
        Call WriteLog("実行時間超過: 処理を中断します", "YoyakuPTN_THU_TRAIN", "タイムアウト")
        GoTo CleanUpDriver
    End If

    ' メインループ - すべてのステーションを1番目候補として試す
    For i = 0 To loopMAX - 1
        ' タイムアウトチェック
        If DateDiff("s", startTime, Now) > maxExecutionTime Then
            Call WriteLog("実行時間超過: 処理を中断します", "YoyakuPTN_THU_TRAIN", "タイムアウト")
            GoTo CleanUpDriver
        End If
        
        ' すべてのステーションをリセット
        Thu_ClearYoyaku PTN, Stallday
        Call WriteLog("--- 1番目候補検索開始: ループ" & (i + 1) & "/" & loopMAX & " ---", "YoyakuPTN_THU_TRAIN", "ループ開始")
        
        ' 使用済みラベルコレクションをリセット
        Set usedLabels = New Collection
        timeoutCounter = 0 ' タイムアウトカウンタリセット
        
        ' 1番目ステーションの選定 - ステーションをインデックス順に試す
        If Not walkSearchCompleted Then
            ' 徒歩検索モード
            ' i番目のステーションを選択
            Call Ascending_StationStation(i)
            stationNo(0) = Asc_ST_Info(0).CarNum
            Call WriteLog("徒歩検索: 1番目候補 = " & stationNo(0), "YoyakuPTN_THU_TRAIN", "徒歩検索開始")
            
            ' このステーションが以前に1番目として使われていないか確認
            Dim alreadyUsed As Boolean
            alreadyUsed = False
            For r = 0 To usedAsFirstCount - 1
                If stationNo(0) = usedAsFirst(r) Then
                    alreadyUsed = True
                    Exit For
                End If
            Next r
            
            If alreadyUsed Then
                Call WriteLog("ステーション" & stationNo(0) & "は既に使用済み: スキップ", "YoyakuPTN_THU_TRAIN", "使用済スキップ")
                GoTo NextFirstStation
            End If
            
            ' A列が空かチェック
            foundFlag = False
            For r = 0 To loopMAX - 1
                With Sheets("Sheet2")
                    If CStr(.Cells(4 + r, "E").Value) = stationNo(0) Then
                        If .Cells(4 + r, "A").Value <> "" Then
                            Call WriteLog("ステーション" & stationNo(0) & "のA列は既に使用済み: スキップ", "YoyakuPTN_THU_TRAIN", "A列使用済")
                            GoTo NextFirstStation
                        End If
                        baseIdx(0) = r: foundFlag = True
                        
                        ' ShopTimeチェック - 徒歩圏内か
                        Dim walkOK As Boolean: walkOK = True
                        Select Case Stallday
                            Case "Stall4day_TueHoliday", "Stall3day_MonTueHoliday"
                                If Asc_ST_Info(0).ShopTime > grpTime Then walkOK = False
                            Case "Stall5day", "Stall4day_ThuHoliday", "Stall4day_FriHoliday", "Stall3day_ThuFriHoliday"
                                If PTN = 2 And Asc_ST_Info(0).ShopTime > grpTime Then walkOK = False
                            Case "Stall4day_MonHoliday", "Stall3day_MonFriHoliday"
                                If PTN = 3 And Asc_ST_Info(0).ShopTime > grpTime Then walkOK = False
                        End Select
                        
                        Call WriteLog("ステーション" & stationNo(0) & " 店舗までの時間: " & Asc_ST_Info(0).ShopTime & "分, 徒歩圏内: " & walkOK, "YoyakuPTN_THU_TRAIN", "徒歩圏内確認")
                        
                        ' 徒歩条件満たさない場合はスキップ
                        If Not walkOK Then
                            Call WriteLog("ステーション" & stationNo(0) & "は徒歩圏外: スキップ", "YoyakuPTN_THU_TRAIN", "徒歩圏外")
                            GoTo NextFirstStation
                        End If
                        
                        ' 使用済みリストに追加
                        usedAsFirst(usedAsFirstCount) = stationNo(0)
                        usedAsFirstCount = usedAsFirstCount + 1
                        Call WriteLog("ステーション" & stationNo(0) & "を1番目に設定", "YoyakuPTN_THU_TRAIN", "1番目設定完了")
                        
                        Exit For
                    End If
                End With
            Next r
            
            If Not foundFlag Then
                Call WriteLog("ステーション" & stationNo(0) & "が見つからない: スキップ", "YoyakuPTN_THU_TRAIN", "ステーション未発見")
                GoTo NextFirstStation
            End If
        Else
            ' 電車検索モード
            ' i番目のステーションを選択
            Call Ascending_StationStation(i)
            stationNo(0) = Asc_ST_Info(0).CarNum
            Call WriteLog("電車検索: 1番目候補 = " & stationNo(0), "YoyakuPTN_THU_TRAIN", "電車検索開始")
            
            ' このステーションが以前に1番目として使われていないか確認
            alreadyUsed = False
            For r = 0 To usedAsFirstCount - 1
                If stationNo(0) = usedAsFirst(r) Then
                    alreadyUsed = True
                    Exit For
                End If
            Next r
            
            If alreadyUsed Then
                Call WriteLog("ステーション" & stationNo(0) & "は既に使用済み: スキップ", "YoyakuPTN_THU_TRAIN", "使用済スキップ")
                GoTo NextFirstStation
            End If
            
            ' A列が空かチェック
            foundFlag = False
            For r = 0 To loopMAX - 1
                With Sheets("Sheet2")
                    If CStr(.Cells(4 + r, "E").Value) = stationNo(0) Then
                        If .Cells(4 + r, "A").Value <> "" Then
                            Call WriteLog("ステーション" & stationNo(0) & "のA列は既に使用済み: スキップ", "YoyakuPTN_THU_TRAIN", "A列使用済")
                            GoTo NextFirstStation
                        End If
                        baseIdx(0) = r: foundFlag = True
                        
                        ' 電車での店舗移動可能性チェック - 特定のStallday/PTN組み合わせのみ
                        Dim trainOK As Boolean: trainOK = True
                        Select Case Stallday
                            Case "Stall4day_TueHoliday", "Stall3day_MonTueHoliday"
                                trainOK = False ' 制限あり - 後でチェック
                            Case "Stall5day", "Stall4day_ThuHoliday", "Stall4day_FriHoliday", "Stall3day_ThuFriHoliday"
                                If PTN = 2 Then trainOK = False ' 制限あり - 後でチェック
                            Case "Stall4day_MonHoliday", "Stall3day_MonFriHoliday"
                                If PTN = 3 Then trainOK = False ' 制限あり - 後でチェック
                        End Select
                        
                        If trainOK Then
                            ' 制限なし - 直接設定
                            Call WriteLog("電車検索: 制限なしで1番目設定: " & stationNo(0), "YoyakuPTN_THU_TRAIN", "制限なし設定")
                            ' 使用済みリストに追加
                            usedAsFirst(usedAsFirstCount) = stationNo(0)
                            usedAsFirstCount = usedAsFirstCount + 1
                        Else
                            ' 制限あり - 電車で店舗まで移動可能かチェック
                            baseAddress = .Cells(4 + baseIdx(0), "G").Value
                            dayToSet = labels(0)
                            Call WriteLog("電車移動チェック: " & baseAddress & "→店舗 " & storeAddress, "YoyakuPTN_THU_TRAIN", "電車移動確認")
                            If GetYahooTransit_UseOptionClick_Safe_WithDriver(driver, baseAddress, dayToSet, 3, grpTime, storeAddress, usedLabels, excludedStations) Then
                                For rr = 0 To loopMAX - 1
                                    If Sheets("Sheet2").Cells(4 + rr, "A").Value = dayToSet Then
                                        stationNo(0) = CStr(Sheets("Sheet2").Cells(4 + rr, "E").Value)
                                        baseIdx(0) = rr: foundFlag = True
                                        Call WriteLog("1番目設定成功: " & stationNo(0) & " → " & dayToSet, "YoyakuPTN_THU_TRAIN", "電車設定完了")
                                        
                                        ' 使用済みリストに追加
                                        usedAsFirst(usedAsFirstCount) = stationNo(0)
                                        usedAsFirstCount = usedAsFirstCount + 1
                                        
                                        Exit For
                                    End If
                                Next rr
                            Else
                                ' タイムアウト発生時の処理
                                If IncrementTimeoutCounter(timeoutCounter, MAX_TIMEOUTS) Then
                                    Call WriteLog("1番目検索でタイムアウト最大回数に達したため次の候補へ", "YoyakuPTN_THU_TRAIN", "タイムアウト上限")
                                    SafeSleep 5000 ' ブラウザ回復のための長めの休憩
                                    GoTo NextFirstStation
                                End If
                                
                                foundFlag = False
                                Call WriteLog("電車移動条件を満たすステーションが見つからない: スキップ", "YoyakuPTN_THU_TRAIN", "電車検索失敗")
                                GoTo NextFirstStation
                            End If
                        End If
                        
                        Exit For
                    End If
                End With
            Next r
            
            If Not foundFlag Then
                Call WriteLog("ステーション" & stationNo(0) & "が見つからない: スキップ", "YoyakuPTN_THU_TRAIN", "ステーション未発見")
                GoTo NextFirstStation
            End If
        End If
        
        excludedStations = stationNo(0)
        Call WriteLog("除外ステーション設定: " & excludedStations, "YoyakuPTN_THU_TRAIN", "除外設定")

        '=== 2番目探索 ===
        Call WriteLog("--- 2番目候補検索開始 ---", "YoyakuPTN_THU_TRAIN", "2番目検索開始")
        timeoutCounter = 0 ' タイムアウトカウンタリセット
        
        Call Ascending_StationStation(baseIdx(0))
        For j = 1 To loopMAX - 1
            ' タイムアウトチェック
            If DateDiff("s", startTime, Now) > maxExecutionTime Then
                Call WriteLog("実行時間超過: 処理を中断します", "YoyakuPTN_THU_TRAIN", "タイムアウト")
                GoTo CleanUpDriver
            End If
            
            stationNo(1) = "": foundFlag = False
            ' 徒歩圏内を全量検索
            For k = 1 To loopMAX - 1
                With Asc_ST_Info(k)
                    If .CarNum <> stationNo(0) And .Time <= grpTime Then
                        Call WriteLog("徒歩検索: 2番目候補 = " & .CarNum & ", 距離時間: " & .Time & "分", "YoyakuPTN_THU_TRAIN", "徒歩検索")
                        candCar = .CarNum
                        For r = 0 To loopMAX - 1
                            If CStr(Sheets("Sheet2").Cells(4 + r, "E").Value) = candCar And Sheets("Sheet2").Cells(4 + r, "A").Value = "" Then
                                stationNo(1) = candCar: baseIdx(1) = r: foundFlag = True
                                Call WriteLog("2番目設定候補: " & stationNo(1) & " (徒歩)", "YoyakuPTN_THU_TRAIN", "徒歩候補発見")
                                Exit For
                            End If
                        Next r
                    End If
                End With
                
                If foundFlag Then Exit For ' 徒歩で見つかったら終了
            Next k
             
            If Not foundFlag Then
                baseAddress = Sheets("Sheet2").Cells(4 + baseIdx(0), "G").Value
                dayToSet = labels(1)
                ' stage2除外を含めた除外リストを構築
                Dim exclude2All As String
                exclude2All = excludedStations
                If excludedStage2 <> "" Then exclude2All = exclude2All & "," & excludedStage2
                Call WriteLog("電車検索: 2番目候補, 除外リスト=" & exclude2All, "YoyakuPTN_THU_TRAIN", "電車検索開始")
                
                ' リクエスト間隔をランダム化
                SafeSleep GetOptimizedWaitTime(800, 1500)
                If GetYahooTransit_UseOptionClick_Safe_WithDriver(driver, baseAddress, dayToSet, 1, grpTime, "", usedLabels, exclude2All) Then
                    For r = 0 To loopMAX - 1
                        If Sheets("Sheet2").Cells(4 + r, "A").Value = dayToSet Then
                            stationNo(1) = CStr(Sheets("Sheet2").Cells(4 + r, "E").Value)
                            baseIdx(1) = r: foundFlag = True
                            Call WriteLog("2番目設定成功: " & stationNo(1) & " → " & dayToSet & " (電車)", "YoyakuPTN_THU_TRAIN", "電車設定完了")
                            Exit For
                        End If
                    Next r
                Else
                    ' タイムアウト発生時の処理
                    If IncrementTimeoutCounter(timeoutCounter, MAX_TIMEOUTS) Then
                        Call WriteLog("2番目検索でタイムアウト最大回数に達したため次の候補へ", "YoyakuPTN_THU_TRAIN", "タイムアウト上限")
                        SafeSleep 5000 ' ブラウザ回復のための長めの休憩
                        GoTo SkipJ_THU
                    End If
                    
                    'バックトラック: 2番目のクリア
                    If stationNo(1) <> "" Then
                        Call WriteLog("バックトラック: 2番目設定をクリア " & stationNo(1), "YoyakuPTN_THU_TRAIN", "バックトラック")
                        For r = 0 To loopMAX - 1
                            If CStr(Sheets("Sheet2").Cells(4 + r, "E").Value) = stationNo(1) Then
                                With Sheets("Sheet2")
                                    .Cells(4 + r, "A").ClearContents
                                    .Range("FS" & (4 + r) & ":FX" & (4 + r)).ClearContents
                                End With: Exit For
                            End If
                        Next r
                        ' ステージ2リストに追加
                        excludedStage2 = IIf(excludedStage2 = "", stationNo(1), excludedStage2 & "," & stationNo(1))
                        Call WriteLog("ステージ2除外リスト追加: " & excludedStage2, "YoyakuPTN_THU_TRAIN", "除外リスト更新")
                        stationNo(1) = ""
                    End If
                    GoTo SkipJ_THU
                End If
            End If
            
            If foundFlag Then
                excludedStations = stationNo(0) & "," & stationNo(1)
                Call WriteLog("除外ステーション更新: " & excludedStations, "YoyakuPTN_THU_TRAIN", "除外設定更新")
                Call WriteLog("--- 3番目候補検索開始 ---", "YoyakuPTN_THU_TRAIN", "3番目検索開始")
                
                ' タイムアウトチェック
                If DateDiff("s", startTime, Now) > maxExecutionTime Then
                    Call WriteLog("実行時間超過: 処理を中断します", "YoyakuPTN_THU_TRAIN", "タイムアウト")
                    GoTo CleanUpDriver
                End If
                
                timeoutCounter = 0 ' タイムアウトカウンタリセット
                
                '=== 3番目探索（修正版：徒歩全量検索→電車検索） ===
                Call Ascending_StationStation(baseIdx(1))
                stationNo(2) = ""
                foundFlag = False
                
                ' 徒歩圏内を全量検索
                For k = 1 To loopMAX - 1
                    With Asc_ST_Info(k)
                        If .CarNum <> stationNo(0) And .CarNum <> stationNo(1) And .Time <= grpTime Then
                            Call WriteLog("徒歩検索: 3番目候補 = " & .CarNum & ", 距離時間: " & .Time & "分", "YoyakuPTN_THU_TRAIN", "徒歩検索")
                            candCar = .CarNum
                            For r = 0 To loopMAX - 1
                                If CStr(Sheets("Sheet2").Cells(4 + r, "E").Value) = candCar And Sheets("Sheet2").Cells(4 + r, "A").Value = "" Then
                                    stationNo(2) = candCar: baseIdx(2) = r: foundFlag = True
                                    Call WriteLog("3番目設定候補: " & stationNo(2) & " (徒歩)", "YoyakuPTN_THU_TRAIN", "徒歩候補発見")
                                    Exit For
                                End If
                            Next r
                        End If
                    End With
                    
                    If foundFlag Then Exit For ' 徒歩で見つかったら終了
                Next k
                
                ' 徒歩で見つからなかった場合のみ電車検索
                If Not foundFlag Then
                    baseAddress = Sheets("Sheet2").Cells(4 + baseIdx(1), "G").Value
                    dayToSet = labels(2)
                    ' stage3除外を含めた除外リストを構築
                    Dim exclude3All As String
                    exclude3All = excludedStations
                    If excludedStage3 <> "" Then exclude3All = exclude3All & "," & excludedStage3
                    Call WriteLog("電車検索: 3番目候補, 除外リスト=" & exclude3All, "YoyakuPTN_THU_TRAIN", "電車検索開始")
                    
                    ' リクエスト間隔をランダム化
                    SafeSleep GetOptimizedWaitTime(800, 1500)
                    If GetYahooTransit_UseOptionClick_Safe_WithDriver(driver, baseAddress, dayToSet, 1, grpTime, "", usedLabels, exclude3All) Then
                        For r = 0 To loopMAX - 1
                            If Sheets("Sheet2").Cells(4 + r, "A").Value = dayToSet Then
                                stationNo(2) = CStr(Sheets("Sheet2").Cells(4 + r, "E").Value)
                                baseIdx(2) = r: foundFlag = True
                                Call WriteLog("3番目設定成功: " & stationNo(2) & " → " & dayToSet & " (電車)", "YoyakuPTN_THU_TRAIN", "電車設定完了")
                                Exit For
                            End If
                        Next r
                    Else
                        ' タイムアウト発生時の処理
                        If IncrementTimeoutCounter(timeoutCounter, MAX_TIMEOUTS) Then
                            Call WriteLog("3番目検索でタイムアウト最大回数に達したため次の候補へ", "YoyakuPTN_THU_TRAIN", "タイムアウト上限")
                            SafeSleep 5000 ' ブラウザ回復のための長めの休憩
                            GoTo SkipK_THU
                        End If
                        
                        'バックトラック: 3番目のクリア
                        If stationNo(2) <> "" Then
                            Call WriteLog("バックトラック: 3番目設定をクリア " & stationNo(2), "YoyakuPTN_THU_TRAIN", "バックトラック")
                            For r = 0 To loopMAX - 1
                                If CStr(Sheets("Sheet2").Cells(4 + r, "E").Value) = stationNo(2) Then
                                    With Sheets("Sheet2")
                                        .Cells(4 + r, "A").ClearContents
                                        .Range("FS" & (4 + r) & ":FX" & (4 + r)).ClearContents
                                    End With: Exit For
                                End If
                            Next r
                            ' ステージ3リストに追加
                            excludedStage3 = IIf(excludedStage3 = "", stationNo(2), excludedStage3 & "," & stationNo(2))
                            Call WriteLog("ステージ3除外リスト追加: " & excludedStage3, "YoyakuPTN_THU_TRAIN", "除外リスト更新")
                            stationNo(2) = ""
                        End If
                        GoTo SkipK_THU
                    End If
                End If
    
                ' 3番目が見つかった場合のみ4番目を探索
                If foundFlag Then
                    excludedStations = stationNo(0) & "," & stationNo(1) & "," & stationNo(2)
                    Call WriteLog("4番目用除外ステーション更新: " & excludedStations, "YoyakuPTN_THU_TRAIN", "除外設定更新")
                    Call WriteLog("--- 4番目候補検索開始 ---", "YoyakuPTN_THU_TRAIN", "4番目検索開始")
                    
                    ' タイムアウトチェック
                    If DateDiff("s", startTime, Now) > maxExecutionTime Then
                        Call WriteLog("実行時間超過: 処理を中断します", "YoyakuPTN_THU_TRAIN", "タイムアウト")
                        GoTo CleanUpDriver
                    End If
                    
                    '=== ４番目探索 ===
                    Call Ascending_StationStation(baseIdx(2))
                    stationNo(3) = ""
                    foundFlag = False
                    timeoutCounter = 0 ' タイムアウトカウンタリセット
    
                    For m = 1 To loopMAX - 1
                        With Asc_ST_Info(m)
                            If .CarNum <> stationNo(0) And .CarNum <> stationNo(1) And .CarNum <> stationNo(2) Then
                                Call WriteLog("4番目候補検討: " & .CarNum & ", 距離時間: " & .Time & "分", "YoyakuPTN_THU_TRAIN", "候補検討")
                                ' 3番目の車両からgrpTime以内の車両のみを条件とする
                                Dim ok As Boolean: ok = (.Time <= grpTime)
                                
                                Call WriteLog("ステーション" & .CarNum & " 3番目からの時間: " & .Time & "分, 条件満たす: " & ok, "YoyakuPTN_THU_TRAIN", "条件確認")
    
                                If ok Then
                                    ' 徒歩OK
                                    Call WriteLog("4番目徒歩OK: " & .CarNum, "YoyakuPTN_THU_TRAIN", "徒歩OK")
                                    candCar = .CarNum
                                    For r = 0 To loopMAX - 1
                                        If CStr(Sheets("Sheet2").Cells(4 + r, "E").Value) = candCar And Sheets("Sheet2").Cells(4 + r, "A").Value = "" Then
                                            stationNo(3) = candCar
                                            Call WriteLog("4番目設定候補: " & stationNo(3) & " (徒歩)", "YoyakuPTN_THU_TRAIN", "徒歩候補発見")
                                            foundFlag = True
                                            foundAll = True
                                            Exit For
                                        End If
                                    Next r
                                Else
                                    ' 徒歩NG → 電車検索 checkFlg=1（修正点）
                                    baseAddress = Sheets("Sheet2").Cells(4 + baseIdx(2), "G").Value
                                    dayToSet = labels(3)
                                    Call WriteLog("電車検索: 4番目候補, ベース=" & baseAddress, "YoyakuPTN_THU_TRAIN", "電車検索開始")
                                    ' リクエスト間隔をランダム化
                                    SafeSleep GetOptimizedWaitTime(800, 1500)
                                    If GetYahooTransit_UseOptionClick_Safe_WithDriver(driver, baseAddress, dayToSet, 1, grpTime, "", usedLabels, excludedStations) Then
                                        For rr = 0 To loopMAX - 1
                                            If Sheets("Sheet2").Cells(4 + rr, "A").Value = dayToSet Then
                                                stationNo(3) = CStr(Sheets("Sheet2").Cells(4 + rr, "E").Value)
                                                Call WriteLog("4番目設定成功: " & stationNo(3) & " → " & dayToSet & " (電車)", "YoyakuPTN_THU_TRAIN", "電車設定完了")
                                                foundFlag = True
                                                foundAll = True
                                                Exit For
                                            End If
                                        Next rr
    
                                    Else
                                        ' タイムアウト発生時の処理
                                        If IncrementTimeoutCounter(timeoutCounter, MAX_TIMEOUTS) Then
                                            Call WriteLog("4番目検索でタイムアウト最大回数に達したため次の候補へ", "YoyakuPTN_THU_TRAIN", "タイムアウト上限")
                                            SafeSleep 5000 ' ブラウザ回復のための長めの休憩
                                            GoTo SkipK_THU
                                        End If
                                        
                                        ' 4番目のバックトラック処理
                                        If stationNo(3) <> "" Then
                                            Call WriteLog("バックトラック: 4番目設定をクリア " & stationNo(3), "YoyakuPTN_THU_TRAIN", "バックトラック")
                                            For r = 0 To loopMAX - 1
                                                If CStr(Sheets("Sheet2").Cells(4 + r, "E").Value) = stationNo(3) Then
                                                    With Sheets("Sheet2")
                                                        .Cells(4 + r, "A").ClearContents
                                                        .Range("FS" & (4 + r) & ":FX" & (4 + r)).ClearContents
                                                    End With
                                                    Exit For
                                                End If
                                            Next r
                                            stationNo(3) = ""
                                            foundFlag = False
                                            foundAll = False
                                        End If
                                    End If
                                End If
    
                                If foundFlag Then Exit For
                            End If
                        End With
                    Next m
                End If
    
SkipK_THU:
                If foundAll Then Exit For
            End If
SkipJ_THU:
        Next j
        If foundAll Then Exit For
NextFirstStation:
        ' すべてのインデックスを試したが徒歩条件で見つからなかった場合
        If i = loopMAX - 1 And Not walkSearchCompleted Then
            i = -1  ' 次のループで0から開始
            walkSearchCompleted = True  ' 電車検索モードに切り替え
            Call WriteLog("すべての候補を徒歩で試行完了: 電車検索モードに切替", "YoyakuPTN_THU_TRAIN", "検索モード切替")
        End If
    Next i

    '=== 書き込み ===
    If Not foundAll Then
        Call WriteLog("全ステーション設定失敗", "YoyakuPTN_THU_TRAIN", "設定失敗")
        GoTo CleanUpDriver
    End If
    
    Call WriteLog("すべてのステーション候補発見: 設定結果を保存", "YoyakuPTN_THU_TRAIN", "書き込み開始")

    ' ラベルカウントとステーション書き込み（重複チェック)
    RET = 0
    With Sheets("Sheet2")
        For j = 0 To 3
            For r = 0 To loopMAX - 1
                If CStr(.Cells(4 + r, "E").Value) = stationNo(j) Then
                    ' 既にこのラベルで設定されているかを確認
                    If .Cells(4 + r, "A").Value = labels(j) Then
                        ' 既に設定済み - カウントのみ
                        Call WriteLog("設定済: " & stationNo(j) & " → " & labels(j), "YoyakuPTN_THU_TRAIN", "設定済確認")
                        RET = RET + 1
                    ElseIf .Cells(4 + r, "A").Value = "" Then
                        ' 空欄なら書き込み
                        .Cells(4 + r, "A").Value = labels(j)
                        Call WriteLog("設定: " & stationNo(j) & " → " & labels(j), "YoyakuPTN_THU_TRAIN", "ラベル設定")
                        RET = RET + 1
                    End If
                    Exit For
                End If
            Next r
        Next j
    End With

    ' driverを閉じる
    If Not driver Is Nothing Then
        On Error Resume Next
        driver.CloseBrowser
        driver.Shutdown
    End If
    
    Call WriteLog("木曜パターン設定 終了: 結果=" & RET, "YoyakuPTN_THU_TRAIN", "処理完了")
    If RET = 4 Then
        YoyakuPTN_THU_TRAIN = RET
        Exit Function
    End If

CleanUpDriver:
    Call WriteLog("クリーンアップ実行", "YoyakuPTN_THU_TRAIN", "クリーンアップ")
    
    ' 強制的にすべてのラベルをクリア
    Thu_ClearYoyaku PTN, Stallday
    
    ' driverを閉じる
    If Not driver Is Nothing Then
        On Error Resume Next
        driver.CloseBrowser
        driver.Shutdown
    End If
    
    Thu_ClearYoyaku PTN, Stallday  ' 2重チェック
    YoyakuPTN_THU_TRAIN = 0
End Function
