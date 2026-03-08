/**
 * AI.gs
 * Gemini API呼び出し共通関数
 */

// ===========================================
// Gemini API共通
// ===========================================

/**
 * Gemini APIを呼び出す
 * @param {string} systemPrompt - システムプロンプト
 * @param {string} userMessage - ユーザーメッセージ
 * @param {string} model - 使用モデル
 * @param {boolean} jsonMode - JSONモードを有効にするか
 * @param {number} temperature - 温度設定（0.0-1.0、デフォルト0.3）
 * @returns {string} AIの応答テキスト
 */
function callGemini(systemPrompt, userMessage, model, jsonMode = false, temperature = 0.3) {
  const apiKey = getGeminiApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: userMessage }]
      }
    ],
    systemInstruction: {
      parts: [{ text: systemPrompt }]
    },
    generationConfig: {
      temperature: temperature
    }
  };

  // JSONモードの場合
  if (jsonMode) {
    payload.generationConfig.responseMimeType = 'application/json';
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    console.error('Gemini API Error:', responseText);
    throw new Error(`Gemini APIエラー (${responseCode}): ${parseGeminiError(responseText)}`);
  }

  const result = JSON.parse(responseText);

  // Geminiのレスポンス形式からテキストを抽出
  if (result.candidates && result.candidates[0] && result.candidates[0].content) {
    return result.candidates[0].content.parts[0].text;
  }

  throw new Error('Gemini APIから有効なレスポンスが返されませんでした');
}

/**
 * Geminiエラーレスポンスをパース
 * @param {string} responseText - レスポンステキスト
 * @returns {string} エラーメッセージ
 */
function parseGeminiError(responseText) {
  try {
    const error = JSON.parse(responseText);
    return error.error?.message || 'Unknown error';
  } catch (e) {
    return responseText.substring(0, 200);
  }
}

// ===========================================
// Calendar用プロンプト
// ===========================================

const CALENDAR_SYSTEM_PROMPT = `あなたは日本語の自然言語テキストから予定情報を1件抽出し、JSONで返すアシスタントです。

## 出力スキーマ（厳守）
{
  "title": "string（必須・空禁止）",
  "start_date": "YYYY-MM-DD",
  "end_date": "YYYY-MM-DD",
  "start_time": "HH:MM または null",
  "end_time": "HH:MM または null",
  "all_day": boolean,
  "memo": "string または null",
  "color_key": "string"
}

## 日時解釈ルール（重要）

### 日付
- 相対日付（明日、来週金曜、3日後 等）は {{TODAY}} を基準に絶対日付へ変換する
- 年が省略されている場合は、{{TODAY}} から見て最も自然な未来日を採用する
- 「来週」「再来週」「来月」等は日本語の自然な暦解釈に従う
- 日をまたぐ時刻範囲（例：22時〜翌1時）は end_date を翌日にする
- 期間予定（「月曜から金曜まで」等）は end_date を適切に設定する
- 単日予定は start_date = end_date

### 時刻
- 明確な時刻（10時、15:30 等）のみ start_time / end_time に入れる
- 曖昧な時刻表現（朝、昼過ぎ、夕方、午後いち、夜 等）は固定時刻に変換しない → null にする
- 時刻が書かれていないだけの予定も、時刻不明なら null にする

### all_day の判定
- 明確に「終日」「1日中」「丸一日」等と言われた場合のみ all_day = true
- 期間予定（複数日にまたがる予定）は all_day = true
- 時刻が不明なだけの予定は all_day = false（時刻未指定と終日は別）
- start_time が null でも、部分的な時間帯の予定なら all_day = false

## title ルール

### 基本
- title は必須。絶対に空にしない
- 予定の主目的・主行為を短い名詞句で表す（目安：2〜20文字）
- 日付・曜日・時刻・助詞（に/で/から 等）は入れない

### title の優先順位
1. 予定の主行為・主目的を採る（「何をするか」）
2. 人物名や日時ではなく、イベント本体を採る
3. 補足や条件は memo に回す

### 禁止ルール
- 次の汎用語だけの title は禁止：「予定」「用事」「タスク」「イベント」「スケジュール」
- ただし内容語を含む複合語は可（例：「通院予定」「面接予定」は可）

### フォールバック
- 抽出困難でも、入力文から内容が分かる短い名詞句を作る
- 最終手段として入力文の先頭20文字以内の要約をtitleにしてよい
- 「（内容不明）」は可。ただし「予定」単独は不可

### 例
- 「歯医者 15時から 保険証持参」→ title: "歯医者", memo: "保険証持参"
- 「明日前田製作所立ち会い」→ title: "前田製作所立ち会い"
- 「明後日、田中さんとランチ」→ title: "田中さんとランチ"
- 「住民税の支払い コンビニでも可」→ title: "住民税支払い", memo: "コンビニでも可"
- 「明日10時に市役所で住民票を取りに行く、マイナンバーカード持参」→ title: "住民票取得", memo: "市役所、マイナンバーカード持参"

## memo ルール

### memo に入れてよい情報
- 持ち物（保険証、マイナンバーカード 等）
- 準備事項
- 注意事項
- 補助的な場所情報（市役所、◯◯駅 等）
- 補助的な相手情報
- 備考として意味がある短文

### memo に入れてはいけない情報
- title の単なる言い換えや繰り返し
- 主予定そのもの（それは title に入れる）
- 日付や時刻の重複記載
- 不要な説明文

### 補足情報がなければ null

## color_key カテゴリ（手段ではなく主目的で分類する）
- health: 医療・健康が主目的（病院、歯医者、検診、薬局、ジム、運動、美容院）
- work: 仕事が主目的（会議、打ち合わせ、出張、面接、仕事関連全般）
- family: 家族・人付き合いが主目的（家族行事、友人、デート、結婚式、お見舞い）
- finance: お金・手続きが主目的（銀行、役所、保険、税金、契約、支払い）
- travel: 旅行・移動自体が主目的（旅行、帰省、引越し、送迎）
- fun: 趣味・娯楽が主目的（映画、コンサート、スポーツ観戦、飲み会、イベント）
- school: 学校・学習が主目的（授業、試験、塾、習い事、PTA、学校行事）
- other: 上記に当てはまらない場合

### color_key 判定の考え方
- 迷ったら「この予定の主目的は何か？」で判断する
- 出張 → 仕事が主目的なら work、旅行自体が目的なら travel
- 学校の保護者会 → 学校行事なら school、家族行事の文脈なら family
- 税理士との打ち合わせ → 税金手続きなら finance、仕事の一環なら work

## 複数予定への対応
- 入力に複数の予定が含まれる場合は、最も主要な1件のみ抽出する
- 他の予定は memo に押し込まない
- 抽出困難な場合は、最初の明確な予定を採用する

## 型ルール（厳守）
- null は文字列 "null" ではなく JSON の null を使う
- all_day は文字列 "true"/"false" ではなく boolean の true/false を使う
- start_date / end_date は必ず YYYY-MM-DD 形式
- start_time / end_time は HH:MM 形式 または null
- 日本の日付形式（◯月◯日）を正しく YYYY-MM-DD に変換する

## 出力前の自己チェック（必須）
1. title が空でないか？禁止語単独でないか？
2. start_date / end_date が YYYY-MM-DD 形式か？
3. start_time / end_time が HH:MM 形式 または null か？
4. all_day が boolean か？
5. null が文字列になっていないか？
6. JSONが壊れていないか？
満たさない場合はルールに従って修正してから出力すること。

## 出力
JSONのみを出力する。説明文は付けない。`;

/**
 * Calendar解析用のシステムプロンプトを取得（日付を埋め込み）
 * @returns {string}
 */
function getCalendarSystemPrompt() {
  const today = getTodayDate();
  return CALENDAR_SYSTEM_PROMPT.replace('{{TODAY}}', today);
}

// ===========================================
// Memo用プロンプト
// ===========================================

const MEMO_SYSTEM_PROMPT = `あなたは音声メモをクリエイティブに整理・まとめ直すAIアシスタントです。

## 目的
ユーザーが話した内容を深く理解し、より価値のある形にまとめ直す

## あなたの4つの役割

### 1. 深い理解
- 話の表面だけでなく、意図や背景を汲み取る
- ユーザーが本当に言いたかったことを理解する
- 文脈から重要なポイントを見抜く

### 2. クリエイティブな整理
- 論理的な構造で情報を整理する
- 必要に応じて箇条書きやカテゴリ分けを使う
- 関連する内容をグループ化する
- 時系列や優先度で並べ替える

### 3. 註釈の追加
- 【補足】として関連情報や背景知識を追記
- 【注意】として気をつけるべき点を明示
- 【ヒント】として役立つアドバイスを提案
- 註釈は控えめに、本当に有用なものだけ追加

### 4. まとめ直し
- 冗長な部分を簡潔にまとめる
- ポイントを明確に抽出する
- 読みやすく、後で見返しやすい形式にする

## フォーマット例

【要点】
・ポイント1
・ポイント2

【詳細】
整理された本文...

【補足】
関連する追加情報...

## 処理ルール
- フィラー（「えー」「あのー」等）は自然に除去
- 元の情報は漏らさず含める（削除しない）
- 追加する註釈は【】で明示して区別する
- 出力はまとめたテキストのみ（説明文は付けない）
- 内容が短い場合は簡潔にまとめる（無理に長くしない）`;


// ===========================================
// API呼び出しラッパー
// ===========================================

/**
 * Calendar用のテキスト解析
 * @param {string} text - ユーザー入力テキスト
 * @returns {Object} 解析結果のJSONオブジェクト
 */
function parseCalendarText(text) {
  const model = getCalendarModel();
  const systemPrompt = getCalendarSystemPrompt();

  const response = callGemini(systemPrompt, text, model, true);

  let parsed;
  try {
    parsed = JSON.parse(response);
  } catch (e) {
    console.error('JSON parse error:', response);
    throw new Error('AIの応答をJSONとしてパースできませんでした');
  }

  // バリデーション（rawTextをフォールバック用に渡す）
  validateCalendarData(parsed, text);

  return parsed;
}

/**
 * Calendar解析結果のバリデーション（フォールバック付き）
 * @param {Object} data - 解析データ
 * @param {string} rawText - 元の入力テキスト（フォールバック用）
 */
function validateCalendarData(data, rawText) {
  // 文字列 "null" を実際の null に修正
  ['start_time', 'end_time', 'memo'].forEach(function(key) {
    if (data[key] === 'null' || data[key] === 'NULL' || data[key] === '') {
      data[key] = null;
    }
  });

  // 文字列 "true"/"false" を boolean に修正
  if (typeof data.all_day === 'string') {
    data.all_day = data.all_day.toLowerCase() === 'true';
  }

  // titleのフォールバック（空禁止）
  if (!data.title || data.title.trim() === '') {
    // 1. memoから先頭20文字を仮タイトルに
    if (data.memo && data.memo.trim()) {
      data.title = data.memo.trim().substring(0, 20);
    }
    // 2. rawTextから先頭20文字を仮タイトルに
    else if (rawText && rawText.trim()) {
      data.title = rawText.trim().substring(0, 20);
    }
    // 3. それでも無理なら「（内容不明）」
    else {
      data.title = '（内容不明）';
    }
  }

  // start_dateのフォールバック（今日）
  if (!data.start_date || !/^\d{4}-\d{2}-\d{2}$/.test(data.start_date)) {
    data.start_date = getTodayDate();
  }

  // end_dateのフォールバック（start_dateと同じ）
  if (!data.end_date || !/^\d{4}-\d{2}-\d{2}$/.test(data.end_date)) {
    data.end_date = data.start_date;
  }

  // end_date が start_date より前なら入れ替え
  if (data.end_date < data.start_date) {
    var tmp = data.start_date;
    data.start_date = data.end_date;
    data.end_date = tmp;
  }

  // 時間のフォーマット検証
  if (data.start_time && !/^\d{2}:\d{2}$/.test(data.start_time)) {
    data.start_time = null;
  }
  if (data.end_time && !/^\d{2}:\d{2}$/.test(data.end_time)) {
    data.end_time = null;
  }

  // all_dayの判定
  // - 明示的にbooleanで返された場合はそのまま使う
  // - 複数日にまたがる予定は all_day = true
  // - 型が不正な場合のみフォールバック
  if (typeof data.all_day !== 'boolean') {
    if (data.start_date !== data.end_date) {
      data.all_day = true;
    } else {
      data.all_day = false;
    }
  }

  // color_keyのバリデーション（許可されたカテゴリのみ）
  const validColorKeys = ['health', 'work', 'family', 'finance', 'travel', 'fun', 'school', 'other'];
  if (!data.color_key || !validColorKeys.includes(data.color_key)) {
    data.color_key = 'other';
  }

  // memoがtitleと同じ内容なら消す（重複防止）
  if (data.memo && data.title && data.memo.trim() === data.title.trim()) {
    data.memo = null;
  }
}

/**
 * Memo用のテキスト整形（クリエイティブAIまとめ）
 * @param {string} rawText - 音声入力の生テキスト
 * @returns {string} 整理・まとめ直されたテキスト
 */
function cleanMemoText(rawText) {
  const model = getMemoModel();

  // クリエイティブな出力のためtemperature 0.7を使用
  const response = callGemini(MEMO_SYSTEM_PROMPT, rawText, model, false, 0.7);

  // 余分な空白を整理
  return response.trim();
}

// ===========================================
// ノート用 3段階AIパイプライン
// ===========================================

// --- Stage 1: 文字起こしクリーンアップ ---
const NOTE_STAGE1_PROMPT = `あなたは文字起こし校正者です。

## 目的
音声認識テキストを、意味を変えずに読みやすく整える。
忠実性を最優先し、要約・再構成・情報追加はしない。

## 処理内容
1. フィラーの除去（「えー」「あー」「えっと」「まあ」「なんか」「あのー」「うーん」等）
2. 音声認識による明らかな誤変換のみ修正する
3. 句読点を適切に補う
4. 文法を最小限だけ整える
5. 不自然な重複や言い直しを整理する
6. 原文の文順は基本的に維持する

## 保持すべき情報（必ず残す）
- 人名・固有名詞
- 日付・時刻
- 数量・金額
- 期限
- 否定表現（〜しない、〜ではない）
- 不確実表現（たぶん、かも、未定、検討中）
- 比較表現（増えた、減った、前回より）

## 禁止事項
- 情報の追加（原文にない内容を足さない）
- 情報の削除（原文にある内容を消さない）
- 意味の言い換え（口語接続を綺麗にしすぎない）
- 構造化（箇条書き化、見出し化をしない）
- 原文にない断定表現への変更
- 文の意味・因果・時制・主語を推測して補わない
- 曖昧な表現は曖昧なまま残す
- 要約しない、段落を再編しない

## 出力
校正後テキストのみを返す（説明文は付けない）`;

// --- Stage 2: 構造化・まとめ ---
const NOTE_STAGE2_PROMPT = `あなたは編集者です。

## 目的
校正済みテキストを読み、内容に最も適した形式に整理する。
原文の情報を過不足なく整理し、推測で補完しない。

## 分類ルール（上から順に判定）
1. 実行すべき項目の列挙が中心 → A: TODO・タスク系
2. 複数人の議論・決定・相談が中心 → B: 会議・相談系
3. 発想・提案・構想・改善案が中心 → C: アイデア・企画系
4. 出来事・状況説明・経過報告が中心 → D: 報告・記録系
5. 短い覚書・断片的メモ・分類困難な短文 → E: メモ・雑記

## AとEの違い
- A: 実行すべき項目が中心（例：「牛乳、洗剤、電球買う」）
- E: 単なる覚書・感想・参照情報（例：「駅前のパン屋、火曜休み」）

## 混在時の優先ルール
- 会議内容の中にタスクが含まれる → Bを優先
- アイデアの中に実行項目がある → Cを優先し、補足にアクションを記載
- 判定が難しい短文 → E

## 出力形式

### A: TODO・タスク系
□ タスク1
□ タスク2
□ タスク3

### B: 会議・相談系
【要点】
・ポイント

【決定事項】
・決まったこと

【アクション】
・動詞で始める（原文に根拠があるもののみ）

【課題】
・未解決事項のみ

### C: アイデア・企画系
【概要】
一言でまとめ

【ポイント】
1. ポイント

【補足】
追加情報

### D: 報告・記録系
【状況】
何が起きたか

【対応】
何をしたか／すべきか

【備考】
補足情報

### E: メモ・雑記
・内容1
・内容2

## 共通ルール
- 原文にない情報を追加しない
- 期限・担当者・数値を推測で補わない
- 原文にない期限は書かない（「来月まで」を「来月末」にしない等）
- 決定事項と推測を混同しない
- アクションは原文に根拠があるものだけ抽出する
- 空になるセクションは無理に作らず省略する
- 各項目は簡潔に1行ずつ記載する
- 内容が短い場合は無理に構造化せず簡潔にまとめる

## 出力
最適な1パターンのみで構造化した結果を返す（説明文は付けない）`;

// --- Stage 3: タイトル生成 ---
const NOTE_STAGE3_PROMPT = `あなたはタイトル生成者です。

## 目的
内容の主題を短く表すタイトルを1つ生成する。
本文の要約ではなく、主題のラベル化を行う。

## ルール
- できるだけ10文字以内にする
- 必要な場合のみ15文字以内まで可
- 内容の本質を表す名詞句にする
- 日付や時刻は含めない
- 「メモ」「ノート」「記録」単独は禁止
- ただし「会議メモ」「買い物メモ」など内容語を含む複合語は可

## 例
- 買い物リストの内容 → 「買い物リスト」
- プロジェクト会議の内容 → 「PJ会議メモ」
- 引越しの準備タスク → 「引越し準備」
- 新サービスのアイデア → 「新サービス案」
- 体調についてのメモ → 「体調メモ」
- 見積提出の確認 → 「見積提出確認」
- 訪問看護シフトの修正 → 「訪看シフト修正」

## 出力
タイトルのみを返す（説明文・引用符・括弧は付けない）`;

/**
 * ノート用3段階AIパイプライン
 * Stage 1: クリーンアップ → Stage 2: 構造化 → Stage 3: タイトル生成
 * @param {string} rawText - 音声入力の生テキスト
 * @param {string} currentTabName - 現在のタブ名（デフォルト名判定用）
 * @param {number} noteId - ノートID（デフォルト名判定用）
 * @returns {Object} { success, cleanedText, structuredText, title, autoTitle }
 */
function processNoteAI(rawText, currentTabName, noteId) {
  try {
    rawText = String(rawText || '').trim();
    if (!rawText) {
      return { success: true, cleanedText: '', structuredText: '', title: '', autoTitle: false };
    }

    const model = getMemoModel();

    // --- Stage 1: クリーンアップ ---
    console.log('Note AI Stage 1: Cleanup');
    const stage1Result = callGemini(NOTE_STAGE1_PROMPT, rawText, model, false, 0.1);
    const cleanedText = (stage1Result || '').trim();

    // --- Stage 2: 構造化 ---
    console.log('Note AI Stage 2: Structure');
    const stage2Result = callGemini(NOTE_STAGE2_PROMPT, cleanedText, model, false, 0.5);
    const structuredText = (stage2Result || '').trim();

    // --- Stage 3: タイトル生成（デフォルト名の場合のみ） ---
    let title = '';
    let autoTitle = false;
    const defaultNames = getDefaultNoteNames();
    const isDefaultName = defaultNames.includes(currentTabName);

    if (isDefaultName && structuredText) {
      console.log('Note AI Stage 3: Title generation');
      const stage3Result = callGemini(NOTE_STAGE3_PROMPT, structuredText, model, false, 0.3);
      title = (stage3Result || '').trim().substring(0, 20); // 安全のため20文字制限
      autoTitle = !!title;
    }

    return {
      success: true,
      cleanedText: cleanedText,
      structuredText: structuredText,
      title: title,
      autoTitle: autoTitle
    };
  } catch (error) {
    console.error('processNoteAI error:', error);
    return {
      success: false,
      cleanedText: '',
      structuredText: '',
      title: '',
      autoTitle: false,
      error: error.message
    };
  }
}

/**
 * 音声文字起こしテキストを整形（後方互換用・カレンダーメモ等で使用）
 * @param {string} rawText - 音声入力の生テキスト
 * @returns {Object} { success: boolean, cleanedText: string, error?: string }
 */
function aiCleanText(rawText) {
  try {
    rawText = String(rawText || '').trim();
    if (!rawText) {
      return { success: true, cleanedText: '' };
    }

    const model = getMemoModel();
    const response = callGemini(NOTE_STAGE1_PROMPT, rawText, model, false, 0.1);

    return {
      success: true,
      cleanedText: (response || '').trim()
    };
  } catch (error) {
    console.error('aiCleanText error:', error);
    return {
      success: false,
      cleanedText: '',
      error: error.message
    };
  }
}