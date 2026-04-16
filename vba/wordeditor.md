# Excel内容をそのままOutlookメール本文にコピー
### メール本文は内部的に Word文書
* WordEditor はその Word.Document オブジェクトなのでWordEditorを使うのはまっとう。
* Dim objMail As Object、Dim doc As Object 宣言は型をあまり考えていない。Dim myInspector As Outlook.Inspector とすることで MailItem と Inspectorで分けることになり設計が明確・将来拡張しやすい
* Dim wdDoc As Word.Document とすることで Word.Document を明示的に扱える
```VisualBasic:Excel内容をそのままOutlookメール本文にコピー
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 待機
Sub ProcessWaitTime(time As Long)
    startTime = [Now()] * 86400000
    Do
        Call Sleep(100) '100ms待機
        DoEvents
    Loop While [Now()] * 86400000 < startTime + time
End Sub

'===========================================================================
' クリップボード検証付き安全貼り付け関数
' doc     : Outlook WordEditor オブジェクト
' waitMs  : Copy後の初回待機時間(ms)。デフォルト200ms推奨
' 戻り値  : 成功=True / 失敗=False
'===========================================================================
Private Function SafePasteToDoc(doc As Object, _
                                Optional waitMs As Long = 200) As Boolean
    Dim retryCount As Integer
    Dim waitTime   As Long
    Dim cbAvail    As Boolean
    
    waitTime = waitMs
    SafePasteToDoc = False
    
    For retryCount = 1 To RETRY_NUM
        ' クリップボードが有効かチェック
        On Error Resume Next
        cbAvail = (UBound(Application.ClipboardFormats) >= 0)
        On Error GoTo 0
        
        If cbAvail Then
            On Error Resume Next
            doc.Windows(1).Selection.Paste
            If Err.Number = 0 Then
                SafePasteToDoc = True
                Exit For
            End If
            Err.Clear
            On Error GoTo 0
        End If
        
        ' 失敗した場合: DoEvents + 待機してリトライ
        DoEvents
        ProcessWaitTime waitTime
        waitTime = waitTime + 100   ' 毎回100msずつ待機を延長
    Next retryCount
End Function


Public Sub CreatedailyMail()

    Dim olApp As Outlook.Application
    ' Outlookが起動していれば取得
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0

    ' 起動していなければ起動
    If olApp Is Nothing Then
        Set olApp = New Outlook.Application
    End If
  
    ' MailItem作成
    Dim mail As Outlook.MailItem
    Set mail = olApp.CreateItem(olMailItem)

    ' Inspector（ウィンドウ）を生成
    mail.Display  ' ← 重要：これがないと WordEditor は Nothing

    mail.To = "saito.mieko@kk.jp.panasonic.com"
    mail.CC = "tanaka.takayuki@jp.panasonic.com"
    mail.Subject = "メール作成テスト"
    

    ' 受信者解決
    Dim resolvedCount As Integer
    resolvedCount = 0
    Dim isResolved As Boolean
    isResolved = False
    Do While resolvedCount < 3 And Not isResolved
        isResolved = mail.Recipients.ResolveAll
        ProcessWaitTime (1000)
        resolvedCount = resolvedCount + 1
    Loop
    
    If Not isResolved Then
        Application.StatusBar = "宛先の名前解決に失敗しました。宛先を確認してください。"  ' ステータスバーに表示
    End If

    ' Inspector を取得
    Dim myInspector As Outlook.Inspector
    Set myInspector = mail.GetInspector
    
    ' WordEditor（本文）を取得
    Dim wdDoc As Word.Document
    Set wdDoc = myInspector.WordEditor
    
    Dim wdRange As Word.Range
    Set wdRange = wdDoc.Range
    wdRange.Collapse wdCollapseEnd
    Dim startPos As Long  ' メール本文への挿入位置を記憶する

    ' Wordとして操作可能
    wdRange.InsertAfter "Hello from Inspector!" & vbCrLf
    wdRange.Collapse wdCollapseEnd ' 次の追記のために文末へ
    
    startPos = wdRange.End  ' 追加前の位置を保存  挿入前の位置を記録
    wdRange.InsertAfter "リンクはこちらから" & vbCrLf  ' 文字を追加
    Set wdRange = wdDoc.Range(startPos, wdDoc.Range.End)  ' 追加された部分だけを Range として取得
    wdRange.Font.Color = vbBlue   ' 文字色を青に変更
    'RGB 指定（推奨）
    'wdRange.Font.Color = RGB(255, 0, 0)   ' 赤
    'wdRange.Font.Color = RGB(0, 128, 0)   ' 緑
    'wdRange.Font.Color = RGB(0, 0, 255)   ' 青
    wdRange.Collapse wdCollapseEnd
End Sub
```

### いろいろ文字色変更、リンク、太字パターン
```VisualBasic:vba
Dim wdRange As Word.Range
Dim startPos As Long

' Excel側でコピー済みの前提
Set wdRange = wdDoc.Range
wdRange.Collapse wdCollapseEnd

' Excel表貼り付け
wdRange.Paste
wdRange.Collapse wdCollapseEnd

' リンク挿入
startPos = wdRange.End
wdRange.InsertAfter "▼ 詳細はこちら"
Set wdRange = wdDoc.Range(startPos, wdDoc.Range.End)

wdDoc.Hyperlinks.Add _
    Anchor:=wdRange, _
    Address:="https://www.example.com"

' 通常文字
wdRange.Collapse wdCollapseEnd
wdRange.InsertAfter vbCrLf & "ご不明点があればご連絡ください。"
```

### グラフを画像として貼る（軽量）
```
ExcelChart.ChartArea.CopyPicture xlScreen, xlPicture
wdRange.Paste
```

### グラフを埋め込み（編集可能）ただし重い
```
ExcelChart.ChartArea.Copy
wdRange.Paste
```
