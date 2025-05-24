// ====== 設定項目 ======
const RECIPIENT_NAME = "Ei-chan"; // メール冒頭の宛名 (例: 山田様、ご担当者様)
// ====================

/**
 * 保存されているAPIキーとメールアドレスを取得する関数。
 * 設定されていない場合はエラーをスローします。
 * @returns {object} APIキーとメールアドレスを含むオブジェクト。例: { apiKey: 'YOUR_API_KEY', emailAddress: 'your_email@example.com' }
 */
function getGeminiApiKeyAndEmailFromProperties() {
  const userProperties = PropertiesService.getUserProperties();
  const apiKey = userProperties.getProperty('GEMINI_API_KEY');
  const emailAddress = userProperties.getProperty('USER_EMAIL_ADDRESS'); // メールアドレスを取得

  if (!apiKey) {
    throw new Error("Gemini APIキーがUser Propertiesに設定されていません。setGeminiApiKeyAndEmail関数を実行してください。");
  }

  if (!emailAddress) {
    throw new Error("メールアドレスがUser Propertiesに設定されていません。setGeminiApiKeyAndEmail関数を実行してください。");
  }

  // APIキーとメールアドレスをオブジェクトとして返す
  return {
    apiKey: apiKey,
    emailAddress: emailAddress
  };
}

function sendDailyBusinessNews() {
  const newsContent = getBusinessNewsFromGemini();
  const mailAddress = getGeminiApiKeyAndEmailFromProperties().emailAddress

  if (newsContent) {
    const subject = "【毎朝お届け】昨日の主要ビジネスニュース速報";
    const body = `
${RECIPIENT_NAME}

昨日の主要なビジネスニュースをGeminiがまとめてお届けします。

---

${newsContent}

---

この情報が皆様のビジネスにお役立ていただければ幸いです。

よろしくお願いいたします。
`;

    MailApp.sendEmail(mailAddress, subject, body);
    Logger.log("メールを送信しました。");
  } else {
    Logger.log("ニュースコンテンツが取得できませんでした。");
  }
}

function getBusinessNewsFromGemini() {
  // ★User PropertiesからAPIキーを取得
  const GEMINI_API_KEY = getGeminiApiKeyAndEmailFromProperties().apiKey;
  const modelName = "gemini-2.5-flash-preview-05-20"; // 使用するGeminiモデル
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${GEMINI_API_KEY}`;

  // 昨日の日付を取得
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy年MM月dd日");

  // Geminiへのプロンプト
  const prompt = `
あなたはビジネスニュース専門のAIアシスタントです。
${formattedYesterday}の主要なビジネスニュースを収集し、以下の要件で簡潔にまとめてください。

* **対象期間:** ${formattedYesterday}の主要なビジネスニュース（日本、米国、欧州市場に焦点を当てる）
* **カテゴリ:** 金融、IT、テクノロジー、製造業、国際経済、M&A、新製品・サービス
* **出力形式:**
    * 読みやすいプレーンテキストで出力
    * 各ニュースのタイトル
    * 簡潔な要約（3行程度）
    * 主要な影響（ビジネスへの影響、市場への影響など）
    * 最大5つの主要なニュースに絞って詳細に記述し、その他注目ニュースは簡潔に箇条書きでまとめること。
* **トーン:** 客観的かつ分析的。
* **その他:** 過度な専門用語は避け、ビジネスパーソンが朝の短時間で把握できるようにまとめること。情報源のURLは不要です。
`;

  const requestBody = {
    contents: [
      {
        parts: [
          {
            text: prompt,
          },
        ],
      },
    ],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody),
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());

    // Geminiからのレスポンスを解析
    if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
      return jsonResponse.candidates[0].content.parts[0].text;
    } else {
      Logger.log("Geminiからのレスポンス形式が不正です。");
      Logger.log(jsonResponse); // デバッグ用にレスポンス全体を出力
      return null;
    }
  } catch (e) {
    Logger.log("Gemini API呼び出し中にエラーが発生しました: " + e.toString());
    return null;
  }
}

function sendWeeklyTechNews() {
  const newsContent = getTechNewsFromGemini();
  const mailAddress = getGeminiApiKeyAndEmailFromProperties().emailAddress;

  if (newsContent) {
    const subject = "【毎週お届け】今週の主要テックニュース";
    const body = `
${RECIPIENT_NAME}

今週の主要なTechニュースをGeminiがまとめてお届けします。

---

${newsContent}

---

この情報が皆様のビジネスにお役立ていただければ幸いです。

よろしくお願いいたします。
`;

    MailApp.sendEmail(mailAddress, subject, body);
    Logger.log("メールを送信しました。");
  } else {
    Logger.log("ニュースコンテンツが取得できませんでした。");
  }
}

function getTechNewsFromGemini() {
  // ★User PropertiesからAPIキーを取得
  const GEMINI_API_KEY = getGeminiApiKeyAndEmailFromProperties().apiKey;
  const modelName = "gemini-2.5-flash-preview-05-20"; // 使用するGeminiモデル
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${GEMINI_API_KEY}`;

  // 昨日の日付を取得 (この部分は削除またはコメントアウト)
  // const yesterday = new Date();
  // yesterday.setDate(yesterday.getDate() - 1);
  // const formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy年MM月dd日");

  // 今週の期間（例：先週の土曜日から金曜日まで）を取得
  const today = new Date();
  const currentDay = today.getDay(); // 0 = 日曜日, 1 = 月曜日, ..., 6 = 土曜日

  // スクリプトが土曜日に実行されることを想定し、先週の土曜日と今週の金曜日を計算
  // 先週の土曜日 (スクリプト実行日の8日前)
  const lastSaturday = new Date(today);
  lastSaturday.setDate(today.getDate() - (currentDay + 7) % 7 - 7); // 土曜なら today.getDate() - 7
  const formattedLastSaturday = Utilities.formatDate(lastSaturday, Session.getScriptTimeZone(), "yyyy年MM月dd日");

  // 今週の金曜日 (スクリプト実行日の1日前)
  const thisFriday = new Date(today);
  thisFriday.setDate(today.getDate() - 1); // 土曜なら today.getDate() - 1
  const formattedThisFriday = Utilities.formatDate(thisFriday, Session.getScriptTimeZone(), "yyyy年MM月dd日");

  // プロンプト内で使用する期間文字列
  const dateRange = `${formattedLastSaturday}から${formattedThisFriday}まで`;

  // Geminiへのプロンプト
  const prompt = `
  あなたはテクノロジーニュース専門のAIアシスタントです。
  ${dateRange}の主要なテックニュースを収集し、以下の要件で簡潔にまとめてください。

  * **対象期間:** ${dateRange}の主要なテックニュース（世界中の主要なトレンド、特に日本、米国、欧州に焦点を当てる）
  * **カテゴリ:** AI、半導体、クラウド、サイバーセキュリティ、Web3、XR/メタバース、宇宙開発、クリーンテック、消費者向け電子機器、スタートアップ投資、規制動向
  * **出力形式:**
      * 週の主要トピックを3〜5個に絞り、それぞれのトピックについて以下の情報を記述する。
      * トピックのタイトル
      * 簡潔な要約（5行程度）
      * その週のトレンドや市場、社会への影響
      * その他、注目すべきニュースは箇条書きで数点まとめること。
  * **トーン:** 客観的、分析的、かつ未来志向。
  * **その他:** 過度な専門用語は避け、ビジネスパーソンが週末に短時間でキャッチアップできるようにまとめること。情報源のURLは不要です。
  `;

  const requestBody = {
    contents: [
      {
        parts: [
          {
            text: prompt,
          },
        ],
      },
    ],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody),
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());

    // Geminiからのレスポンスを解析
    if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
      return jsonResponse.candidates[0].content.parts[0].text;
    } else {
      Logger.log("Geminiからのレスポンス形式が不正です。");
      Logger.log(jsonResponse); // デバッグ用にレスポンス全体を出力
      return null;
    }
  } catch (e) {
    Logger.log("Gemini API呼び出し中にエラーが発生しました: " + e.toString());
    return null;
  }
}

function sendDailyBusinessEnglishWords() {
  const englishWordsContent = getDailyBusinessEnglishWordsFromGemini();
  const mailAddress = getGeminiApiKeyAndEmailFromProperties().emailAddress;

  if (englishWordsContent) {
    const subject = "【毎朝お届け】ビジネス英単語・熟語10選";
    const body = `
${RECIPIENT_NAME}

いつもお世話になっております。
本日は、昨日のビジネスニュースからピックアップした、ビジネスで役立つC1レベルの英単語・熟語10選をお届けします。
毎日新しい語彙を学ぶことで、あなたのビジネス英語スキルを向上させましょう。

---

${englishWordsContent}

---

この情報が皆様の英語学習にお役立ていただければ幸いです。

よろしくお願いいたします。
`;

    MailApp.sendEmail(mailAddress, subject, body);
    Logger.log("ビジネス英単語メールを送信しました。");
  } else {
    Logger.log("ビジネス英単語コンテンツが取得できませんでした。");
  }
}

function getDailyBusinessEnglishWordsFromGemini() {
  const GEMINI_API_KEY = getGeminiApiKeyAndEmailFromProperties().apiKey;
  const modelName = "gemini-2.5-flash-preview-05-20"; // 使用するGeminiモデル
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${GEMINI_API_KEY}`;

  // 昨日の日付を取得（ビジネスニュースと同じロジック）
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy年MM月dd日");

  // ここが最も重要なプロンプトの部分です
  const prompt = `
あなたは英語学習者向けのAI講師です。
以下の手順で、ビジネス英語のC1レベルの英単語・熟語を10個生成してください。

**手順:**
1.  まず、${formattedYesterday}の主要なビジネスニュース（特に日本、米国、欧州市場の動向、金融、テクノロジー、経済全般）を1つピックアップしてください。
2.  ピックアップしたビジネスニュースの簡単な概要（1〜2文）を記述してください。
3.  そのニュースの中で使われていた、またはそのニュースに関連するビジネスで頻出するC1レベルの英単語・熟語を10個選定してください。
4.  選定した各単語・熟語について、以下の形式で情報を提供してください。

    * **英単語/熟語:**
    * **意味:** （日本語で簡潔に）
    * **ニュースからの例文:** （ピックアップしたビジネスニュースから、その単語・熟語が実際に使われていたような文を生成してください。現実のニュースからの引用のように見せてください。）
    * **ビジネスシーンでの使用例:** （会議、メール、プレゼンなどでどのように使えるか、具体的な例文を1つ）

**出力形式の注意点:**
* 各単語は番号付きリストで表示し、間に改行を挟むなどして見やすくしてください。
* 全体的にプロフェッショナルで教育的なトーンで記述してください。
* 毎日異なる単語が選ばれるように、多様なニュースや文脈から選ぶことを意識してください。
`;

  const requestBody = {
    contents: [
      {
        parts: [
          {
            text: prompt,
          },
        ],
      },
    ],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody),
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
      return jsonResponse.candidates[0].content.parts[0].text;
    } else {
      Logger.log("Geminiからのレスポンス形式が不正です。");
      Logger.log(jsonResponse);
      return null;
    }
  } catch (e) {
    Logger.log("Gemini API呼び出し中にエラーが発生しました: " + e.toString());
    return null;
  }
}