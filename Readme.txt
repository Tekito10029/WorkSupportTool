WorkSupportTool 関数処理概要
============================

このファイルは、WorkSupportTool の主要な C++ 実装ファイルで定義されている関数が、どのような処理を担当しているかをまとめたものです。
対象ファイルは Main.cpp、SearchToolPage.cpp、PrintToolPage.cpp です。


1. Main.cpp
-----------
メインウィンドウ、上部タブ、検索ページと印刷ページの切り替えを管理します。

| 関数 | 処理概要 |
| --- | --- |
| ApplyModernControlTheme | 指定された Windows コントロールに Explorer テーマを適用し、標準コントロールの見た目を統一します。 |
| FillRoundRect | 指定矩形を角丸で塗りつぶし、枠線も描画します。タブ背景描画の共通処理です。 |
| DrawModernTab | メインタブをオーナードローで描画します。選択中のタブは白背景と青い下線で強調します。 |
| GetPageRect | メインウィンドウのクライアント領域から、タブ下に配置するページ領域を計算します。 |
| LayoutMain | メインウィンドウのリサイズに合わせて、タブ、検索ページ、印刷ページを再配置します。 |
| ApplyMainTab | 現在選択されているタブに応じて検索ページ/印刷ページの表示を切り替えます。印刷タブ表示時には検索結果のファイルパスを印刷ページへ渡します。 |
| MainWndProc | メインウィンドウのウィンドウプロシージャです。作成、サイズ変更、タブ切り替え、背景描画、オーナードロー、終了時のリソース解放などを処理します。 |
| wWinMain | アプリケーションのエントリーポイントです。COM と共通コントロールを初期化し、検索/印刷ページのウィンドウクラスを登録してメインウィンドウを作成し、メッセージループを実行します。 |


2. SearchToolPage.cpp
---------------------
検索ツールページを構成します。検索対象フォルダー、検索期間、対象拡張子、フォルダー除外、ファイル名除外、検索結果一覧、CSV 出力、設定保存/復元、独自 UI 描画を担当します。

| 関数 | 処理概要 |
| --- | --- |
| ToLower | 文字列を小文字化して返します。比較やフィルター用です。 |
| Trim | 文字列前後の空白、タブ、改行を取り除きます。 |
| EllipsizePathRight | 長いパスを右側省略表示用に短縮します。 |
| GetRootsFromListBox | 検索ルート一覧リストボックスからパス文字列だけを取得します。 |
| BuildRootDisplayText | 検索ルートの表示文字列を作成します。無効なルートには無効状態を示す接頭辞を付けます。 |
| ParseRootDisplayText | 表示文字列から実際のルートパスと有効/無効状態を取り出します。 |
| GetRootEntriesFromListBox | ルート一覧から、パスと有効状態を持つ RootEntry の配列を作ります。 |
| SetRootEntriesToListBox | RootEntry 配列をリストボックスへ反映します。 |
| GetEnabledRootsFromListBox | 有効状態の検索ルートだけを取得します。 |
| ToggleSelectedRootEnabled | 選択中ルートの有効/無効を切り替え、リスト表示を更新します。 |
| RootExistsInListBox | 指定パスがルート一覧に既に存在するか確認します。 |
| AddRootToListBoxDedup | 重複を避けながらルートパスをリストへ追加します。 |
| MoveRootItem | ルート一覧の項目を上下に移動します。 |
| SplitRootsText | INI などに保存された複数ルート文字列を分割して配列化します。 |
| JoinRootsForIni | ルート一覧を INI 保存用の区切り文字列に変換します。 |
| GetWindowTextWStr | ウィンドウ/コントロールのテキストを std::wstring として取得します。 |
| GetExeDir | 実行ファイルが置かれているフォルダーを取得します。 |
| NormalizePath | パスを絶対パス化し、可能であれば weakly_canonical で正規化します。 |
| IsPathSeparator | 文字がパス区切り文字か判定します。 |
| HasPathPrefixLow | 小文字化済みのパス同士で、ディレクトリ境界を考慮した前方一致判定を行います。 |
| FormatLocalTime | system_clock の時刻をローカル時刻の「YYYY-MM-DD HH:MM:SS」文字列にします。 |
| GetLocalDayRangePastNDays | 今日または過去 N 日のローカル日付範囲を開始/終了時刻として計算します。 |
| IsWithinLocalRange | 指定時刻が検索対象期間内に含まれるか判定します。 |
| GetMode | 検索期間モード（今日、過去 N 日、期間指定）をコンボボックスから取得します。 |
| GetDaysFromEdit | 過去 N 日入力欄の値を読み取り、1～3650 日に補正します。 |
| LocalMidnightFromDate | SYSTEMTIME の日付をローカル時刻の 0:00 の time_point に変換します。 |
| GetActiveDateRange | 現在の検索期間設定から、実際に検索に使う開始/終了時刻を決定します。 |
| Progress_SetMarquee | 進捗バーのマーキー表示を開始/停止し、停止時は空表示へ戻します。 |
| FileTimeToSysClock | Windows FILETIME を std::chrono::system_clock::time_point に変換します。 |
| GetTimeBase | 日時基準（更新日時、作成日時、どちらか）をコンボボックスから取得します。 |
| TimeBaseText | 日時基準の列見出し/表示用テキストを返します。 |
| GetFileTimesSysClock | ファイルの更新日時と作成日時を取得し、system_clock に変換します。 |
| ExtChecked | 指定拡張子のチェックボックスが選択されているか判定します。 |
| AnyExtSelected | 対象拡張子チェックボックスのいずれかが選択されているか判定します。 |
| UseTargetExtensionFilter | 拡張子フィルターを使うべきか判定します。 |
| SetTargetExtensionSectionEnabled | 拡張子チェックボックス群の有効/無効を切り替えます。 |
| IsTargetExcelFile | パスの拡張子が選択中の Excel 対象拡張子に該当するか判定します。 |
| WideToUtf8 | UTF-16 の wstring を UTF-8 文字列へ変換します。 |
| Utf8ToWide | UTF-8 文字列を UTF-16 の wstring へ変換します。 |
| ReadTextFileUtf8 | UTF-8 テキストファイルを読み込み、wstring として返します。 |
| WriteTextFileUtf8Bom | UTF-8 BOM 付きテキストファイルを書き出します。 |
| PickFolder | フォルダー選択ダイアログを表示し、選択パスを取得します。 |
| PickOpenFile | ファイルを開くダイアログを表示し、選択ファイルを取得します。 |
| PickSaveFile | 名前を付けて保存ダイアログを表示し、保存先パスを取得します。 |
| FolderRuleToDisplay | フォルダー除外ルールをリスト表示用文字列に変換します。 |
| RefreshExcludeListBox | フォルダー除外ルール一覧をリストボックスへ再描画します。 |
| IsExcludedDir | 正規化済みフォルダーが除外ルールに該当するか判定します。 |
| AddExcludeDirPrefix | フォルダー除外（ディレクトリ前方一致）ルールを追加します。 |
| AddOrUpdateExcludePatternOrSubstring | ワイルドカードまたは部分一致のフォルダー除外ルールを追加/更新します。 |
| LoadExcludesFromFile | 除外フォルダー設定ファイルを読み込み、除外ルールへ反映します。 |
| SaveExcludesToFile | 現在のフォルダー除外ルールをファイルへ保存します。 |
| UpdateExcludeDirPrefixAt | 指定位置のフォルダー除外ルールを新しいパスで更新します。 |
| CommitExcludeEditIfNeeded | 除外パターン入力欄の内容があれば、追加/更新処理として確定します。 |
| FileNamePatternExistsExcept | 指定したファイル名除外パターンが、指定インデックス以外に重複しているか確認します。 |
| CommitFileNameEditIfNeeded | ファイル名除外入力欄の内容を追加/更新として確定します。 |
| ExclEditProc | フォルダー除外入力欄のサブクラスプロシージャです。Enter で確定、Esc でキャンセルなどを扱います。 |
| FNameEditProc | ファイル名除外入力欄のサブクラスプロシージャです。Enter で確定、Esc でキャンセルなどを扱います。 |
| RebuildFileNameExcludeCache | ファイル名除外パターンの小文字化キャッシュを再構築します。 |
| RefreshFileNameListBox | ファイル名除外一覧をリストボックスへ反映します。 |
| IsExcludedByFileName | ファイル名が除外パターンに該当するか判定します。 |
| LoadFileNameExcludesFromFile | ファイル名除外設定ファイルを読み込みます。 |
| SaveFileNameExcludesToFile | ファイル名除外パターンをファイルへ保存します。 |
| GetFilterLow | 検索結果フィルター入力欄の文字列を小文字化して取得します。 |
| HitMatchesFilterLow | 検索結果 1 件がフィルター文字列に一致するか判定します。 |
| UpdateExportButtonEnabled | 検索結果の有無に応じて CSV 出力ボタンの有効/無効を更新します。 |
| CsvEscape | CSV セル用にダブルクォートのエスケープと囲み処理を行います。 |
| ExportResultsCsv | 検索結果を CSV ファイルへ出力します。 |
| InitListViewColumns | 検索結果リストビューの列を初期化します。 |
| SetListViewTimeHeader | 日時基準に合わせて検索結果リストの日時列見出しを更新します。 |
| ClearResultsUI | 検索結果 UI と内部結果リストをクリアします。 |
| AddResultToUI | 検索ヒットを内部リストへ追加し、現在のフィルターに合う場合は画面にも追加します。 |
| RebuildListViewFromResults | フィルター変更などに合わせ、内部結果からリストビューを再構築します。 |
| SortResults | 指定列と昇順/降順で検索結果を並べ替えます。 |
| ContainsQuestionMark | 文字列に '?' が含まれるか判定します。ワイルドカード判定などに使います。 |
| CreateUtf16LeBomFile | UTF-16LE BOM 付きの空ファイルを作成します。 |
| IsUnicodeIniFile | INI ファイルが Unicode として扱えるか確認します。 |
| MakeUniqueBackupPath | 既存ファイルのバックアップ名が重複しないよう、一意なパスを生成します。 |
| IniReadStrFrom | 指定 INI ファイルから文字列値を読み取ります。 |
| IniReadIntFrom | 指定 INI ファイルから整数値を読み取ります。 |
| GetLocalAppDataExcelTodayDir | LocalAppData 配下の検索ツール用設定フォルダーを取得/作成します。 |
| InitPaths | 設定ファイルや標準除外ファイルなど、アプリで使うパスを初期化します。 |
| EnsureUnicodeIniWithMigration | INI ファイルを Unicode 化し、必要に応じて既存設定を移行/バックアップします。 |
| IniWriteStr | 現在の INI ファイルへ文字列値を書き込みます。 |
| IniWriteInt | 現在の INI ファイルへ整数値を書き込みます。 |
| IniReadStr | 現在の INI ファイルから文字列値を読み取ります。 |
| IniReadInt | 現在の INI ファイルから整数値を読み取ります。 |
| UpdateUiEnableStates | 検索モード、除外有効フラグ、検索中状態などに応じて各 UI の有効/無効を更新します。 |
| EnsureThemeBrushes | 背景やカード用のブラシを必要に応じて作成します。 |
| ApplyModernControlTheme | 指定コントロールへ Explorer テーマを適用します。 |
| EnableModernOwnerDrawButton | ボタンをオーナードロー形式に変更します。 |
| ApplyModernListBox | リストボックスにテーマを適用し、項目高さを設定します。 |
| UpdateHotControl | マウスホバー中コントロールの状態を更新し、再描画します。 |
| StartHoverTracking | マウスリーブイベントを受け取るための追跡を開始します。 |
| ModernButtonHoverProc | ボタンのホバー状態を管理するサブクラスプロシージャです。 |
| EnableButtonHoverHighlight | ボタンへホバー強調用サブクラスを設定します。 |
| DrawModernCheckBoxFace | チェックボックスを独自デザインで描画します。 |
| RedrawModernCheckBoxNow | チェックボックスを即時再描画します。 |
| ToggleModernCheckBox | チェックボックスのチェック状態を切り替えて再描画します。 |
| CallCheckBoxDefaultWithoutNativePaint | ネイティブ描画を抑えつつ、チェックボックス既定処理を呼び出します。 |
| ModernCheckBoxProc | チェックボックスのクリック、キー操作、ホバー、フォーカスを処理するサブクラスプロシージャです。 |
| ApplyModernCheckBox | チェックボックスに独自描画用スタイルとサブクラスを設定します。 |
| DrawModernComboBoxFace | コンボボックスの選択欄を独自デザインで描画します。 |
| ModernComboBoxProc | コンボボックスのホバーや再描画を扱うサブクラスプロシージャです。 |
| FormatDateText | SYSTEMTIME を日付表示文字列へ変換します。 |
| DateForPicker | 対象の日付ピッカーが開始日/終了日のどちらかを判定し、対応する日付変数を返します。 |
| UpdateModernDatePickerText | 独自日付ピッカーの表示テキストを更新します。 |
| IsLeapYear | うるう年か判定します。 |
| DaysInMonth | 指定年月の日数を返します。 |
| DayOfWeek | 指定年月日の曜日を計算します。 |
| OffsetCalendarMonth | 独自カレンダーの表示月を前後に移動します。 |
| CloseModernCalendarPopup | 表示中の独自カレンダーポップアップを閉じます。 |
| DrawModernDatePickerFace | 日付ピッカーの見た目を独自描画します。 |
| PaintModernCalendarPopup | 独自カレンダーポップアップ全体を描画します。 |
| HitTestModernCalendarDay | カレンダー上の座標からクリックされた日付を判定します。 |
| ModernCalendarPopupProc | カレンダーポップアップの描画、クリック、キー操作、終了を処理します。 |
| ShowModernCalendarPopup | 指定した日付ピッカー用にカレンダーポップアップを表示します。 |
| ApplyModernComboBox | コンボボックスへオーナードロー、項目高さ、サブクラスを設定します。 |
| ApplyModernResultsListView | 検索結果リストビューへテーマ、行高さ、フル行選択などを設定します。 |
| ApplyModernDatePickerTheme | 日付ピッカーへ独自描画用サブクラスを設定します。 |
| IsPrimaryButtonId | 指定 ID が主要操作ボタンか判定します。 |
| IsDangerButtonId | 指定 ID が危険操作（削除など）ボタンか判定します。 |
| IsCsvButtonId | 指定 ID が CSV 出力ボタンか判定します。 |
| DrawRoundedRect | 角丸矩形を塗りつぶし色と枠線色で描画します。 |
| DrawModernButton | ボタンを用途、ホバー、押下、無効状態に応じて独自描画します。 |
| DrawModernListBox | リストボックス項目を選択状態や交互背景を考慮して独自描画します。 |
| HandleResultsHeaderCustomDraw | 検索結果リストのヘッダー部分を独自描画します。 |
| DrawModernComboBox | コンボボックスの選択欄/ドロップダウン項目を独自描画します。 |
| HandleResultsCustomDraw | 検索結果リストビューの行やサブ項目を独自色で描画します。 |
| DrawModernTab | 左ペインのタブを独自描画します。 |
| DrawCard | カード状の背景領域を描画します。 |
| PaintSearchBackground | 検索ページ全体の背景とカード領域を描画します。 |
| ApplyLeftTabVisibility | 左ペインタブの選択状態に応じて、検索条件 UI と除外設定 UI の表示を切り替えます。 |
| SetLeftTab | 左ペインタブを切り替え、必要なら設定へ保存します。 |
| LoadSettings | INI から検索条件、除外設定、表示状態などを読み込んで UI へ反映します。 |
| SaveSettings | 現在の検索条件、除外設定、表示状態などを INI へ保存します。 |
| SetSearchingUi | 検索中/停止中の状態に応じてボタン、進捗、ステータス表示を切り替えます。 |
| DoLayout | 検索ページ内の全コントロールをウィンドウサイズに合わせて配置します。 |
| CopyTextToClipboard | 指定文字列をクリップボードへコピーします。 |
| ShowResultsContextMenu | 検索結果一覧の右クリックメニューを表示し、パスコピーやフォルダーを開く操作を提供します。 |
| ShowRootsContextMenu | 検索ルート一覧の右クリックメニューを表示します。 |
| ShowExcludesContextMenu | フォルダー除外一覧の右クリックメニューを表示します。 |
| ShowFNameContextMenu | ファイル名除外一覧の右クリックメニューを表示します。 |
| SearchThreadProc | バックグラウンドで検索を実行します。対象ルートを再帰走査し、拡張子、日時範囲、除外条件に一致した Excel ファイルを UI スレッドへ通知します。 |
| GetModeTextForStatus | 現在の検索モードをステータス表示用の文字列に変換します。 |
| StartSearch | 入力チェック、検索条件作成、結果クリア、UI 更新を行い、検索スレッドを開始します。 |
| StopSearch | 検索中断フラグを立て、検索停止を要求します。 |
| WndProc | 検索ページのウィンドウプロシージャです。コントロール生成、コマンド、通知、描画、検索スレッドからのメッセージ、設定保存、リソース解放を処理します。 |
| SearchToolPage_GetResultPaths | 検索結果からファイルパス一覧を抽出して返します。印刷ページへの受け渡しに使います。 |
| RegisterSearchToolPageClass | 検索ページ用ウィンドウクラスを登録します。 |
| CreateSearchToolPage | 検索ページ子ウィンドウを作成します。 |
| SearchToolPage_SetVisible | 検索ページの表示/非表示を切り替えます。 |
| SearchToolPage_Resize | 検索ページを指定領域へ移動/リサイズします。 |


3. PrintToolPage.cpp
--------------------
印刷ツールページを構成します。検索結果や手動追加ファイルを印刷対象にし、Excel COM を使って指定シートをプレビューまたは印刷します。プリンター、用紙、部数、シート設定、ログ表示、独自 UI 描画を担当します。

| 関数 | 処理概要 |
| --- | --- |
| EnsureThemeBrushes | ページ背景、カード、入力欄用のブラシを必要に応じて作成します。 |
| ApplyModernControlTheme | 指定コントロールへ Explorer テーマを適用します。 |
| EnableModernOwnerDrawButton | ボタンをオーナードロー形式に変更します。 |
| ApplyModernListBox | リストボックスにテーマを適用し、項目高さを設定します。 |
| UpdateHotControl | ホバー中コントロールを更新し、変更前後のコントロールを再描画します。 |
| StartHoverTracking | マウスリーブイベントを受け取るための追跡を開始します。 |
| ModernButtonHoverProc | ボタンのホバー状態を管理するサブクラスプロシージャです。 |
| EnableButtonHoverHighlight | ボタンへホバー強調用サブクラスを設定します。 |
| DrawModernCheckBoxFace | チェックボックスを独自デザインで描画します。 |
| RedrawModernCheckBoxNow | チェックボックスを即時再描画します。 |
| ToggleModernCheckBox | チェックボックスのチェック状態を反転し、通知と再描画を行います。 |
| CallCheckBoxDefaultWithoutNativePaint | ネイティブ描画を抑えつつ、チェックボックス既定処理を呼び出します。 |
| ModernCheckBoxProc | チェックボックスのクリック、キー、ホバー、フォーカスを処理するサブクラスプロシージャです。 |
| ApplyModernCheckBox | チェックボックスへ独自描画用スタイルとサブクラスを設定します。 |
| DrawModernComboBoxFace | コンボボックス選択欄を独自デザインで描画します。 |
| ModernComboBoxProc | コンボボックスのホバー状態と再描画を扱うサブクラスプロシージャです。 |
| ApplyModernComboBox | コンボボックスへオーナードロー、項目高さ、サブクラスを設定します。 |
| DrawRoundedRect | 角丸矩形を塗りつぶし色と枠線色で描画します。 |
| DrawModernButton | ボタンを用途、ホバー、押下、無効状態に応じて独自描画します。 |
| DrawModernListBox | リストボックス項目を選択状態や交互背景を考慮して独自描画します。 |
| DrawCard | カード状の白い背景と枠線を描画します。 |
| DrawModernComboBox | コンボボックスの選択欄/一覧項目を独自描画します。 |
| PaintPageBackground | 印刷ページ全体の背景とカード領域を描画します。 |
| GetDefaultPrinterName | Windows の既定プリンター名を取得します。 |
| SetDefaultPrinterName | Windows の既定プリンターを指定名に変更します。 |
| GetExeDir | 実行ファイルが置かれているフォルダーを取得します。 |
| GetLocalAppDataPrintToolDir | LocalAppData 配下の印刷ツール用設定フォルダーを取得/作成します。 |
| GetIniPath | 印刷ツール設定 INI ファイルのパスを返します。 |
| AddLog | ログ欄の末尾に 1 行追記します。 |
| RefreshFileList | 内部の印刷対象ファイル一覧をリストボックスへ反映します。 |
| GetCurrentFileLineIndex | 印刷対象リストで現在選択されている行番号を取得します。 |
| UpdateRemoveTargetLabel | 削除対象ラベルに現在選択中のファイル名を表示します。 |
| RemoveCurrentFileLine | 選択中の印刷対象ファイルを一覧から削除し、ログを出します。 |
| Trim | 文字列前後の空白、タブ、改行を除去します。 |
| SplitSheetNames | シート名入力を改行、カンマ、セミコロン、読点で分割し、空要素を除外します。 |
| GetCopies | 部数入力欄から部数を取得し、1～99 に補正します。 |
| GetCheck | チェックボックスがチェックされているか返します。 |
| PickFiles | Excel ファイル選択ダイアログを表示し、単一または複数選択されたファイルパスを取得します。 |
| AutoWrap | IDispatch のプロパティ取得/設定やメソッド呼び出しを共通化した COM 呼び出しヘルパーです。 |
| GetDispatchProperty | COM オブジェクトの dispatch 型プロパティを取得します。 |
| CallMethod | COM オブジェクトの引数なしメソッドを呼び出します。 |
| CallMethodN | COM オブジェクトの複数引数メソッドを呼び出します。 |
| GetWorksheetByName | Excel Worksheets コレクションから指定名のワークシートを取得します。 |
| SetBoolProperty | COM オブジェクトの bool プロパティを設定します。 |
| SetStringProperty | COM オブジェクトの文字列プロパティを設定します。 |
| SetLongProperty | COM オブジェクトの long プロパティを設定します。 |
| OpenWorkbook | Excel Workbooks コレクションから指定ファイルを開き、Workbook オブジェクトを取得します。 |
| FreeDevMode | 保持しているプリンター DEVMODE メモリを解放します。 |
| CreateDefaultDevModeForPrinter | 指定プリンターの既定 DEVMODE を取得して保持します。 |
| RefreshPaperCombo | 現在選択中プリンターの用紙一覧を取得し、用紙コンボボックスへ反映します。 |
| RefreshPrinterCombo | 利用可能なプリンター一覧を取得し、プリンターコンボボックスへ反映します。 |
| GetSelectedPaperCode | 用紙コンボボックスで選択されている用紙コードを返します。 |
| GetPrinterNameFromDevNames | PRINTDLG の DEVNAMES からプリンター名を取り出します。 |
| BuildActivePrinterStringFromPrinterName | Excel の ActivePrinter に渡すプリンター指定文字列を作成します。 |
| GetPaperCodeFromDevMode | DEVMODE から用紙コードを取得します。 |
| GetTrayCodeFromDevMode | DEVMODE から給紙トレイコードを取得します。 |
| GetColorModeFromDevMode | DEVMODE からカラー/モノクロ設定を取得します。 |
| GetCopiesFromDevModeOrDefault | DEVMODE の部数設定を取得し、なければ指定した既定値を返します。 |
| ShowRuntimePrintDialog | 印刷直前の Windows 印刷ダイアログを表示し、プリンターや部数などの設定を反映します。 |
| SaveSheetSettings | シート名、部数、プレビュー設定を INI に保存します。 |
| LoadSheetSettings | INI からシート名、部数、プレビュー設定を読み込みます。 |
| ShowPrinterProperties | 選択中プリンターのプロパティ画面を表示し、設定変更後に用紙一覧を更新します。 |
| BuildActivePrinterString | 現在選択中のプリンターから Excel ActivePrinter 用文字列を作成します。 |
| ShowWorksheetPrintPreview | 指定ワークシートの印刷プレビューを Excel COM で表示します。 |
| PrintWorksheet | 指定ワークシートを指定部数、プレビュー有無、プリンターで印刷します。 |
| PrintOneBook | 対象ブックを 1 冊開き、指定シートに対して印刷設定を適用し、プレビューまたは印刷を実行します。印刷中は既定プリンターを一時変更し、終了後に戻します。 |
| ShowPrintSetupDialog | 印刷設定ダイアログを表示し、プリンター、用紙、部数などを更新します。 |
| DoPrint | 入力チェック後、Excel COM を起動して印刷対象ファイルを順に処理します。ログ出力、エラー表示、UI の印刷中状態制御も行います。 |
| LayoutPage | 印刷ページ内の全コントロールをウィンドウサイズに合わせて配置します。 |
| PrintToolPageWndProc | 印刷ページのウィンドウプロシージャです。UI 作成、ボタン操作、プリンター選択、描画、サイズ変更、設定保存、リソース解放を処理します。 |
| PrintToolPage_SetFiles | 検索結果などから渡されたファイル一覧を印刷対象へ設定し、重複排除とソート後に表示を更新します。 |
| RegisterPrintToolPageClass | 印刷ページ用ウィンドウクラスを登録します。 |
| CreatePrintToolPage | 印刷ページ子ウィンドウを作成します。 |
| PrintToolPage_SetVisible | 印刷ページの表示/非表示を切り替えます。 |
| PrintToolPage_Resize | 印刷ページを指定領域へ移動/リサイズします。 |
