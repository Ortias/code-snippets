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

' WordEditor で 文字＋色＋太字＋下線を指定
' 黒・通常文字・下線なしがデフォルトで必要なときだけ指定する
Private Sub AppendFormattedText( _
    ByVal wdDoc As Word.Document, _
    ByRef wdRange As Word.Range, _
    ByVal text As String, _
    Optional ByVal fontColor As Long = wdColorAutomatic, _
    Optional ByVal isBold As Boolean = False, _
    Optional ByVal underlineStyle As WdUnderline = wdUnderlineNone _
)
    Dim startPos As Long
    Dim fmtRange As Word.Range

    ' 追記開始位置を保存
    wdRange.Collapse wdCollapseEnd
    startPos = wdRange.End

    ' 文字を追加
    wdRange.InsertAfter text

    ' 追加した部分だけを取得
    Set fmtRange = wdDoc.Range(startPos, wdRange.End)

    ' 書式設定
    With fmtRange.Font
        .Color = fontColor
        .Bold = isBold
        .Underline = underlineStyle
    End With

    ' 次の追記のために文末へ
    wdRange.Collapse wdCollapseEnd
End Sub

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
    mail.Display

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
        Application.StatusBar = "宛先の名前解決に失敗しました。宛先を確認してください。"
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
    AppendFormattedText wdDoc, wdRange, "Hello from Inspector!" & vbCrLf

    AppendFormattedText wdDoc, wdRange, "リンクはこちらから", vbBlue, False, wdUnderlineNone
    
    startPos = wdRange.End
    wdRange.InsertAfter "▼ 詳細はこちら"
    Set wdRange = wdDoc.Range(startPos, wdDoc.Range.End)
    wdDoc.Hyperlinks.Add _
        Anchor:=wdRange, _
        Address:="https://www.example.com"
    
    AppendFormattedText wdDoc, wdRange, vbCrLf & "続けてみる。" & vbCrLf

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

' Wordとして操作可能
wdRange.InsertAfter "Hello from Inspector!" & vbCrLf
wdRange.Collapse wdCollapseEnd ' 次の追記のために文末へ

startPos = wdRange.End  ' 追加前の位置を保存
wdRange.InsertAfter "リンクはこちらから"
Set wdRange = wdDoc.Range(startPos, wdDoc.Range.End)
wdRange.Font.Color = RGB(0, 0, 255)
wdRange.Collapse wdCollapseEnd

startPos = wdRange.End
wdRange.InsertAfter "▼ 詳細はこちら"
Set wdRange = wdDoc.Range(startPos, wdDoc.Range.End)

wdDoc.Hyperlinks.Add _
    Anchor:=wdRange, _
    Address:="https://www.example.com"

startPos = wdRange.End
wdRange.InsertAfter vbCrLf & "続けてみる。" & vbCrLf
Set wdRange = wdDoc.Range(startPos, wdDoc.Range.End)
wdRange.Font.Color = wdColorAutomatic ' 既定色（黒）
wdRange.Collapse wdCollapseEnd


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
