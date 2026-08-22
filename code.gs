const SHEET_CARDS = 'Cards';
const SHEET_LOG = 'ScanLog';
const SHEET_DECKS = 'DeckSummary';

/**
 * Webアプリを開いたときに表示する画面
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getDecks') {
    return jsonResponse({
      success: true,
      decks: getDecks()
    });
  }

  if (e && e.parameter && e.parameter.action === 'searchCards') {
    const criteria = {
      query: e.parameter.query || '',
      colors: e.parameter.colors ? e.parameter.colors.split(',') : [],
      manaValue: e.parameter.manaValue || '',
      minPriceMin: e.parameter.minPriceMin || '',
      minPriceMax: e.parameter.minPriceMax || '',
      status: e.parameter.status || ''
    };
    return jsonResponse({
      success: true,
      cards: typeof searchCards === 'function' ? searchCards(criteria) : []
    });
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('MTG Card Scanner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 初回セットアップ
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Cardsシート
  let cards = ss.getSheetByName(SHEET_CARDS);

  if (!cards) {
    cards = ss.insertSheet(SHEET_CARDS);
  }

  if (cards.getLastRow() === 0) {
    cards.appendRow([
      'card_name',
      'card_english_name',
      'color',
      'type',
      'mana_value',
      'set_code',
      'collector_number',
      'language',
      'foil',
      'status',
      'deck_id',
      'count',
      'updated_at'
    ]);
  }

  // ScanLogシート
  let log = ss.getSheetByName(SHEET_LOG);

  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
  }

  if (log.getLastRow() === 0) {
    log.appendRow([
      'timestamp',
      'card_name',
      'card_english_name',
      'set_code',
      'collector_number',
      'foil',
      'status',
      'deck_id',
      'method'
    ]);
  }

  // DeckSummaryシート
  let decks = ss.getSheetByName(SHEET_DECKS) || ss.getSheetByName('deck_summary');

  if (!decks) {
    decks = ss.insertSheet(SHEET_DECKS);
  }

  if (decks.getLastRow() === 0) {
    decks.appendRow([
      'deck_id',
      'deck_name'
    ]);
  }

  return 'セットアップ完了';
}

/**
 * DeckSummaryシートからデッキ一覧を取得
 */
function getDecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DECKS) || ss.getSheetByName('deck_summary');

  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(SHEET_DECKS) || ss.getSheetByName('deck_summary');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];
  const deckIdCol = headers.indexOf('deck_id');
  const deckNameCol = headers.indexOf('deck_name');

  if (deckIdCol === -1 || deckNameCol === -1) {
    return [];
  }

  const decks = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const deckId = String(row[deckIdCol] != null ? row[deckIdCol] : '').trim();
    const deckName = String(row[deckNameCol] != null ? row[deckNameCol] : '').trim();
    if (deckId || deckName) {
      decks.push({
        deck_id: deckId,
        deck_name: deckName || deckId
      });
    }
  }

  return decks;
}

/**
 * Scryfallでカード名を検索
 */
function searchScryfallCard(cardName) {

  if (!cardName) {
    throw new Error('カード名がありません');
  }

  const url =
    'https://api.scryfall.com/cards/named?fuzzy=' +
    encodeURIComponent(cardName);

  const response = UrlFetchApp.fetch(url, {
    method: 'get',

    headers: {
      'User-Agent': 'MTG-Card-Scanner/1.0 contact: mtg-card-scanner',
      'Accept': 'application/json;q=0.9,*/*;q=0.8'
    },

    muteHttpExceptions: true
  });

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  // 成功
  if (status === 200) {

    const card =
      JSON.parse(body);

    return convertScryfallCard(card);
  }

  // エラー内容を詳しく表示
  let errorMessage = body;

  try {
    const errorJson =
      JSON.parse(body);

    errorMessage =
      errorJson.details ||
      errorJson.warnings ||
      errorJson.error ||
      body;

  } catch (e) {
    // JSONでなければそのまま使う
  }

  throw new Error(
    'Scryfall APIエラー\n' +
    'HTTP Status: ' + status + '\n' +
    'Message: ' + errorMessage
  );
}

/**
 * Cardsシートのヘッダーから列番号を取得
 */
function getColumnMap(sheet) {

  const headers =
    sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];

  const map = {};

  headers.forEach(function(header, index) {

    if (header) {
      map[String(header).trim()] = index + 1;
    }

  });

  return map;
}


/**
 * Cardsシートへカードを登録
 *
 * 列順はヘッダー名から自動判定するため、
 * Cardsシートの列を自由に並び替え可能。
 */
function registerCard(card) {

  if (!card) {
    throw new Error('カード情報がありません');
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const cardsSheet =
    ss.getSheetByName(SHEET_CARDS);

  const logSheet =
    ss.getSheetByName(SHEET_LOG);

  if (!cardsSheet || !logSheet) {
    throw new Error(
      'CardsまたはScanLogシートがありません'
    );
  }


  /*
   * ヘッダーから列番号を取得
   */
  let columns =
    getColumnMap(cardsSheet);


  /*
   * 必要なカラムを確認（存在しない場合は自動で末尾に追加）
   */
  const expectedColumns = [
    'card_name',
    'card_english_name',
    'color',
    'type',
    'mana_value',
    'set_code',
    'collector_number',
    'language',
    'foil',
    'status',
    'deck_id',
    'count',
    'updated_at'
  ];

  let addedNewColumn = false;
  expectedColumns.forEach(function(col) {
    if (!columns[col]) {
      const lastCol = cardsSheet.getLastColumn();
      cardsSheet.getRange(1, lastCol + 1).setValue(col);
      columns[col] = lastCol + 1;
      addedNewColumn = true;
    }
  });

  if (addedNewColumn) {
    columns = getColumnMap(cardsSheet);
  }


  const now =
    new Date();


  /*
   * status, deck_id の決定
   * status取りうる値: deck, case, storage, selling
   * statusが"deck"の時だけdeck_idに値が入る
   */
  const allowedStatuses = ['deck', 'case', 'storage', 'selling'];
  let status = card.status;
  if (allowedStatuses.indexOf(status) === -1) {
    status = 'storage';
  }
  const deckId = (status === 'deck') ? String(card.deck_id || '').trim() : '';
  const registeredCardName = cleanJapaneseCardName(card.card_name);


  /*
   * Cardsシートのデータを取得
   */
  const values =
    cardsSheet
      .getDataRange()
      .getValues();


  let existingRow = -1;


  /*
   * 既存カードを検索
   * 判定キー: セットコード + コレクター番号 + 言語 + foil
   * （セット情報がない場合は カード名 + foil）
   */
  const targetSetCode = String(card.set_code || '').trim().toLowerCase();
  const targetCollectorNumber = String(card.collector_number || '').trim().toLowerCase();
  const targetLanguage = String(card.language || '').trim().toLowerCase();
  const targetFoil = String(card.foil || false).toLowerCase();
  const targetCardName = String(registeredCardName || '').trim().toLowerCase();
  const targetCardEnglishName = String(card.card_english_name || '').trim().toLowerCase();

  for (
    let i = 1;
    i < values.length;
    i++
  ) {
    const row = values[i];
    const rowSetCode = columns['set_code'] ? String(row[columns['set_code'] - 1] || '').trim().toLowerCase() : '';
    const rowCollectorNumber = columns['collector_number'] ? String(row[columns['collector_number'] - 1] || '').trim().toLowerCase() : '';
    const rowLanguage = columns['language'] ? String(row[columns['language'] - 1] || '').trim().toLowerCase() : '';
    const rowFoil = columns['foil'] ? String(row[columns['foil'] - 1] || false).toLowerCase() : 'false';
    const rowCardName = columns['card_name'] ? String(row[columns['card_name'] - 1] || '').trim().toLowerCase() : '';
    const rowCardEnglishName = columns['card_english_name'] ? String(row[columns['card_english_name'] - 1] || '').trim().toLowerCase() : '';

    if (targetSetCode && targetCollectorNumber) {
      if (
        rowSetCode === targetSetCode &&
        rowCollectorNumber === targetCollectorNumber &&
        (!targetLanguage || !rowLanguage || rowLanguage === targetLanguage) &&
        rowFoil === targetFoil
      ) {
        existingRow = i + 1;
        break;
      }
    } else if (targetCardName || targetCardEnglishName) {
      if (
        (rowCardName === targetCardName || (targetCardEnglishName && rowCardEnglishName === targetCardEnglishName)) &&
        rowFoil === targetFoil
      ) {
        existingRow = i + 1;
        break;
      }
    }
  }


  /*
   * 登録するカード情報
   */
  const cardData = {

    card_name:
      registeredCardName,

    card_english_name:
      card.card_english_name || '',

    set_code:
      card.set_code || '',

    collector_number:
      card.collector_number || '',

    language:
      card.language || '',

    foil:
      card.foil || false,

    status:
      status,

    deck_id:
      deckId,

    color:
      card.color || '',

    type:
      card.type || '',

    mana_value:
      card.mana_value != null
        ? card.mana_value
        : ''

  };


  /*
   * 既存カード
   */
  if (existingRow !== -1) {

    /*
     * countを取得
     */
    const countColumn =
      columns['count'];

    const countCell =
      cardsSheet.getRange(
        existingRow,
        countColumn
      );

    const currentCount =
      Number(
        countCell.getValue()
      ) || 0;


    /*
     * 枚数 +1
     */
    countCell.setValue(
      currentCount + 1
    );


    /*
     * カード情報を更新
     */
    Object.keys(cardData)
      .forEach(function(key) {

        const column =
          columns[key];

        if (!column) {
          return;
        }

        const value =
          cardData[key];

        // deck_id は status が deck 以外の場合空文字にするため空文字でも上書き
        if (key === 'deck_id') {
          cardsSheet
            .getRange(
              existingRow,
              column
            )
            .setValue(value);
          return;
        }

        // status は常に反映
        if (key === 'status') {
          cardsSheet
            .getRange(
              existingRow,
              column
            )
            .setValue(value);
          return;
        }

        /*
         * 空値で既存データを消さない
         */
        if (
          value === '' ||
          value === null ||
          value === undefined
        ) {
          return;
        }

        cardsSheet
          .getRange(
            existingRow,
            column
          )
          .setValue(value);

      });


    /*
     * updated_at
     */
    cardsSheet
      .getRange(
        existingRow,
        columns['updated_at']
      )
      .setValue(now);

  }


  /*
   * 新規カード
   */
  else {

    /*
     * 新しい行を作成
     */
    const newRow =
      cardsSheet.getLastRow() + 1;


    /*
     * カード情報を書き込む
     */
    Object.keys(cardData)
      .forEach(function(key) {

        const column =
          columns[key];

        if (!column) {
          return;
        }

        cardsSheet
          .getRange(
            newRow,
            column
          )
          .setValue(
            cardData[key]
          );

      });


    /*
     * count
     */
    cardsSheet
      .getRange(
        newRow,
        columns['count']
      )
      .setValue(1);


    /*
     * updated_at
     */
    cardsSheet
      .getRange(
        newRow,
        columns['updated_at']
      )
      .setValue(now);

  }


  /*
   * ScanLogには毎回記録
   */
  logSheet.appendRow([

    now,

    registeredCardName,

    card.card_english_name || '',

    card.set_code || '',

    card.collector_number || '',

    card.foil || false,

    status,

    deckId,

    'Scryfall'

  ]);


  return {

    success: true,

    message:
      (registeredCardName || card.card_english_name || 'カード') +
      ' を登録しました'

  };

}


function testScryfall() {

  const result =

    searchScryfallCard('Lightning Bolt');

  Logger.log(result);

}

/**
 * OCRで得たカード名からScryfall検索
 */
function findCardByName(cardName) {

  if (!cardName) {
    throw new Error(
      'カード名がありません'
    );
  }

  return searchScryfallCard(
    cardName
  );

}

/**
 * Gemini APIの接続テスト
 */
function testGemini() {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error(
      "GEMINI_API_KEY が設定されていません"
    );
  }

  const model =
    "gemini-3.5-flash-lite";

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/"
    + model
    + ":generateContent?key="
    + encodeURIComponent(apiKey);


  const payload = {

    contents: [

      {
        parts: [

          {
            text:
              "これはテストです。日本語で「Gemini接続成功」とだけ回答してください。"
          }

        ]
      }

    ]

  };


  const response =
    UrlFetchApp.fetch(

      url,

      {
        method: "post",

        contentType:
          "application/json",

        payload:
          JSON.stringify(payload),

        muteHttpExceptions:
          true

      }

    );


  const code =
    response.getResponseCode();

  const body =
    response.getContentText();


  console.log(
    "HTTP Status:",
    code
  );

  console.log(
    "Response:",
    body
  );


  if (code < 200 || code >= 300) {

    throw new Error(
      "Gemini APIエラー: " +
      code +
      " " +
      body
    );

  }


  const json =
    JSON.parse(body);


  const text =
    json.candidates?.[0]
      ?.content?.parts?.[0]
      ?.text;


  console.log(
    "Gemini:",
    text
  );


  return text;

}

/**
 * MTGカード画像をGeminiで判定
 *
 * @param {string} base64Image JPEG画像のBase64
 * @return {Object} Geminiの判定結果
 */
function identifyMtgCard(base64Image) {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error(
      "GEMINI_API_KEY が設定されていません"
    );
  }


  const model =
    "gemini-3.5-flash-lite";


  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/"
    + model
    + ":generateContent";


  const prompt = `
あなたはMagic: The Gatheringのカード識別専門家です。

添付された画像に写っているMTGカードを特定してください。

カードの印刷物を注意深く確認し、以下の情報を読み取ってください。

- カード名
- セットコード
- セット名
- コレクター番号
- 言語

特にカード名、セットコード、コレクター番号を重視してください。

画像から読み取れない情報は推測せず null にしてください。

カードがMTGカードではない場合も明示してください。

以下のJSON形式だけで回答してください。

{
  "is_mtg_card": true,
  "card_name": "カード名",
  "set_code": "セットコード",
  "collector_number": "コレクター番号",
  "language": "en",
  "confidence": 0.0
}

confidenceは0.0～1.0の範囲で、
画像からカードを特定できる確信度を表してください。
`;


  const payload = {

    contents: [

      {
        parts: [

          {
            text: prompt
          },

          {
            inline_data: {

              mime_type:
                "image/jpeg",

              data:
                base64Image

            }

          }

        ]

      }

    ],

    generationConfig: {

      temperature: 0.1,

      responseMimeType:
        "application/json"

    }

  };


  const response =
    UrlFetchApp.fetch(

      url,

      {

        method: "post",

        contentType:
          "application/json",

        headers: {

          "x-goog-api-key":
            apiKey

        },

        payload:
          JSON.stringify(payload),

        muteHttpExceptions:
          true

      }

    );


  const status =
    response.getResponseCode();

  const body =
    response.getContentText();


  console.log(
    "Gemini HTTP:",
    status
  );

  console.log(
    "Gemini response:",
    body
  );


  if (
    status < 200 ||
    status >= 300
  ) {

    throw new Error(
      "Gemini APIエラー: " +
      status +
      " " +
      body
    );

  }

  const json =
    callGeminiWithFallback(payload);

  const text =
    json
      .candidates?.[0]
      ?.content?.parts?.[0]
      ?.text;


  if (!text) {

    throw new Error(
      "Geminiから回答が取得できませんでした"
    );

  }


const result =
  JSON.parse(text);

console.log(
  "Geminiカード判定結果:",
  JSON.stringify(
    result,
    null,
    2
  )
);


// MTGカードではない
if (result.is_mtg_card === false) {
  return result;
}


// カード名が取れなかった
if (!result.card_name) {
  throw new Error(
    "Geminiからカード名を取得できませんでした"
  );
}


// Scryfallで正式なカード情報を取得
/*
 * Scryfallで正式なカード情報を取得
 *
 * 優先順位：
 *
 * 1. セットコード＋コレクター番号
 * 2. カード名
 */

let scryfallCard;


/*
 * セットコード＋コレクター番号が
 * Geminiから取得できた場合
 */
if (
  result.set_code &&
  result.collector_number
) {

  console.log(
    "Scryfall完全照合:",
    result.set_code,
    result.collector_number
  );


  try {

    scryfallCard =
      searchScryfallBySetAndCollector(
        result.set_code,
        result.collector_number
      );


  } catch (error) {

    /*
     * 完全照合できなかった場合は
     * カード名検索へフォールバック
     */

    console.warn(
      "完全照合失敗。カード名検索へフォールバック:",
      error.message
    );


    scryfallCard =
      searchScryfallCard(
        result.card_name
      );

  }

}


/*
 * セットコードまたは
 * コレクター番号が読み取れなかった場合
 */
else {

  console.log(
    "セット情報が取得できなかったためカード名検索"
  );


  scryfallCard =
    searchScryfallCard(
      result.card_name
    );

}


/*
 * Scryfall照合結果
 */

console.log(
  "Scryfall照合結果:",
  JSON.stringify(
    scryfallCard,
    null,
    2
  )
);


/*
 * Geminiの情報と
 * Scryfallの正式情報を統合
 */
return {

  is_mtg_card: true,

  card_name:
    scryfallCard.card_name,

  card_english_name:
    scryfallCard.card_english_name,

  set_code:
    scryfallCard.set_code,

  collector_number:
    scryfallCard.collector_number,

  language:
    scryfallCard.language,

  color:
    scryfallCard.color,

  type:
    scryfallCard.type,

  mana_value:
    scryfallCard.mana_value,

  confidence:
    result.confidence
};

}

/**
 * GitHub PagesからのPOSTを受け取る
 */
function doPost(e) {
  try {
    const data =
      JSON.parse(
        e.postData.contents
      );
    // デッキ一覧取得
    if (data.action === "getDecks") {
      const decks =
        getDecks();
      return jsonResponse({
        success: true,
        decks: decks
      });
    }

    // カード検索
    if (data.action === "searchCards") {
      const cards =
        typeof searchCards === 'function'
          ? searchCards(data.criteria || {})
          : [];
      return jsonResponse({
        success: true,
        cards: cards
      });
    }

    // 登録
    if (data.action === "register") {
      const result =
        registerCard(
          data.card
        );
      return jsonResponse(
        result
      );
    }

    // カード判定
    if (data.image) {
      const result =
        identifyMtgCard(
          data.image
        );

      return jsonResponse({
        success: true,
        result: result
      });
    }

    throw new Error(
      "不正なリクエストです"
    );

  } catch (error) {
    console.error(error);
    return jsonResponse({
      success: false,
      error:
        error.message
    });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * セットコード＋コレクター番号で
 * Scryfallからカードを完全照合する
 *
 * 1. まず日本語版エンドポイント (/cards/:code/:number/ja) を試行
 * 2. 日本語版がなければ通常エンドポイント (/cards/:code/:number) を取得
 */
function searchScryfallBySetAndCollector(
  setCode,
  collectorNumber
) {

  if (!setCode || !collectorNumber) {
    throw new Error(
      'セットコードまたはコレクター番号がありません'
    );
  }

  const cleanSetCode =
    String(setCode).toLowerCase().trim();

  const cleanCollectorNumber =
    String(collectorNumber).trim();

  // 1. まず日本語版直接取得を試行
  const jaUrl =
    'https://api.scryfall.com/cards/' +
    encodeURIComponent(cleanSetCode) +
    '/' +
    encodeURIComponent(cleanCollectorNumber) +
    '/ja';

  try {
    const jaResponse =
      UrlFetchApp.fetch(jaUrl, {
        method: 'get',
        headers: {
          'User-Agent':
            'MTG-Card-Scanner/1.0 contact: mtg-card-scanner',
          'Accept':
            'application/json;q=0.9,*/*;q=0.8'
        },
        muteHttpExceptions: true
      });

    if (jaResponse.getResponseCode() === 200) {
      const jaCard =
        JSON.parse(jaResponse.getContentText());
      return convertScryfallCard(jaCard);
    }
  } catch (e) {
    console.warn('日本語版直接取得スキップ: ' + e.message);
  }

  // 2. 日本語版が直接取れない場合はデフォルト印刷を取得
  const url =
    'https://api.scryfall.com/cards/' +
    encodeURIComponent(cleanSetCode) +
    '/' +
    encodeURIComponent(cleanCollectorNumber);

  const response =
    UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'User-Agent':
          'MTG-Card-Scanner/1.0 contact: mtg-card-scanner',
        'Accept':
          'application/json;q=0.9,*/*;q=0.8'
      },
      muteHttpExceptions: true
    });

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  if (status !== 200) {
    throw new Error(
      'Scryfall完全照合に失敗しました\n' +
      'HTTP Status: ' +
      status +
      '\n' +
      body
    );
  }

  const card =
    JSON.parse(body);

  return convertScryfallCard(card);
}

const SHEET_PRICE_LOG = 'PriceLog';


function setupPriceLog() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  let sheet =
    ss.getSheetByName(SHEET_PRICE_LOG);

  if (!sheet) {
    sheet =
      ss.insertSheet(SHEET_PRICE_LOG);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'timestamp',
      'card_name',
      'card_english_name',
      'set_code',
      'collector_number',
      'min_price',
      'avg_price'
    ]);
  }

  return 'PriceLogセットアップ完了';
}

/**
 * Scryfallのカード情報から色を取得
 *
 * 例：
 * W       → 白
 * U       → 青
 * W,U     → 白,青
 * 無色     → 無色
 */
function getCardColor(card) {

  if (!card) {
    return '';
  }

  // 通常はこちら
  let colors = card.colors || [];

  // colorsが取れない場合はcolor_identityを利用
  if (colors.length === 0 && card.color_identity) {
    colors = card.color_identity;
  }

  const colorMap = {
    W: '白',
    U: '青',
    B: '黒',
    R: '赤',
    G: '緑'
  };

  if (colors.length === 0) {
    return '無色';
  }

  return colors
    .map(function(color) {
      return colorMap[color] || color;
    })
    .join(',');
}

/**
 * Scryfallのtype_lineからカード種別を取得
 *
 * 複数タイプを保持する。
 *
 * 例：
 * Artifact Creature
 * → アーティファクト,クリーチャー
 *
 * Legendary Creature — Human
 * → クリーチャー
 *
 * Basic Land — Island
 * → 土地
 */
function getCardType(card) {

  if (!card) {
    return '';
  }

  const typeLine =
    String(card.type_line || '');

  if (!typeLine) {
    return '';
  }

  const foundTypes = [];

  const typeMap = {
    'Creature': 'クリーチャー',
    'Instant': 'インスタント',
    'Sorcery': 'ソーサリー',
    'Artifact': 'アーティファクト',
    'Enchantment': 'エンチャント',
    'Planeswalker': 'プレインズウォーカー',
    'Land': '土地',
    'Battle': 'バトル',
    'Tribal': '部族',

    'クリーチャー': 'クリーチャー',
    'インスタント': 'インスタント',
    'ソーサリー': 'ソーサリー',
    'アーティファクト': 'アーティファクト',
    'エンチャント': 'エンチャント',
    'プレインズウォーカー': 'プレインズウォーカー',
    '土地': '土地',
    'バトル': 'バトル',
    '部族': '部族'
  };

  Object.keys(typeMap).forEach(function(type) {

    if (typeLine.indexOf(type) !== -1) {

      const japaneseType =
        typeMap[type];

      if (
        foundTypes.indexOf(japaneseType) === -1
      ) {
        foundTypes.push(japaneseType);
      }

    }

  });

  return foundTypes.join(',');
}

/**
 * カード名から「//」以降を除去し、第1面（表面）の名称のみを抽出
 * 例: "Delver of Secrets // Insectile Aberration" -> "Delver of Secrets"
 *     "秘密を掘り下げる者 // 昆虫の逸脱者" -> "秘密を掘り下げる者"
 */
function cleanCardName(name) {
  if (!name) {
    return '';
  }
  const str = String(name).trim();
  const slashIdx = str.indexOf('//');
  if (slashIdx !== -1) {
    return str.substring(0, slashIdx).trim();
  }
  return str;
}

/**
 * 登録用の日本語カード名から丸括弧と注記を除去する。
 * 半角・全角の括弧に対応し、英語名は変更しない。
 */
function cleanJapaneseCardName(name) {
  const cardName = String(name || '').trim();

  if (!/[\u3040-\u30ff\u3400-\u9fff]/.test(cardName)) {
    return cardName;
  }

  let cleaned = cardName;
  let previous;
  do {
    previous = cleaned;
    cleaned = cleaned.replace(/\s*[\(（][^\(\)（）]*[\)）]\s*/g, ' ');
  } while (cleaned !== previous);

  return cleaned.replace(/\s+/g, ' ').trim();
}

/**
 * Scryfall検索クエリを実行し、日本語カードオブジェクトから日本語名を抽出
 * @param {string} query
 * @return {string}
 */
function searchJapanesePrinting(query) {
  try {
    const url =
      'https://api.scryfall.com/cards/search?q=' +
      encodeURIComponent(query);

    const response =
      UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          'User-Agent':
            'MTG-Card-Scanner/1.0 contact: mtg-card-scanner',
          'Accept':
            'application/json;q=0.9,*/*;q=0.8'
        },
        muteHttpExceptions: true
      });

    if (response.getResponseCode() === 200) {
      const data =
        JSON.parse(response.getContentText());

      if (data.data && data.data.length > 0) {
        for (let i = 0; i < data.data.length; i++) {
          const cardObj = data.data[i];

          // 1. 単面カードの printed_name
          if (cardObj.printed_name) {
            return cleanCardName(cardObj.printed_name);
          }

          // 2. 多面カード（両面・分割・当事者等）の 第1面の printed_name
          if (cardObj.card_faces && Array.isArray(cardObj.card_faces) && cardObj.card_faces.length > 0) {
            const firstFace = cardObj.card_faces[0];
            if (firstFace.printed_name) {
              return cleanCardName(firstFace.printed_name);
            }
            if (firstFace.name) {
              return cleanCardName(firstFace.name);
            }
          }

          // 3. printed_nameがなく、langがjaの場合のname
          if (cardObj.lang === 'ja' && cardObj.name) {
            return cleanCardName(cardObj.name);
          }
        }
      }
    }
  } catch (e) {
    console.warn('Scryfall日本語版検索失敗 (' + query + '): ' + e.message);
  }
  return '';
}

/**
 * Scryfallから日本語名を取得・補完する
 * @param {string} oracleId Scryfallのoracle_id
 * @param {string} englishName 英語名
 * @return {string} 日本語名
 */
function fetchJapaneseName(oracleId, englishName) {
  if (!oracleId && !englishName) {
    return '';
  }

  // 1. oracle_id による検索（言語に依存せず確実に同一カードを特定できるため最優先）
  if (oracleId) {
    const jaName =
      searchJapanesePrinting('oracleid:' + oracleId + ' lang:ja include:extras');
    if (jaName) {
      return cleanCardName(jaName);
    }
  }

  // 2. 英語オラクル名の完全一致による検索
  if (englishName) {
    const cleanName = cleanCardName(englishName).replace(/"/g, '');
    const jaName =
      searchJapanesePrinting('!"' + cleanName + '" lang:ja include:extras');
    if (jaName) {
      return cleanCardName(jaName);
    }
  }

  return '';
}

/**
 * Scryfallから英語名を取得・補完する
 * @param {string} oracleId Scryfallのoracle_id
 * @param {string} cardName 日本語名など
 * @return {string} 英語名
 */
function fetchEnglishName(oracleId, cardName) {
  if (!oracleId && !cardName) {
    return '';
  }

  try {
    let url = '';
    if (oracleId) {
      url =
        'https://api.scryfall.com/cards/search?q=' +
        encodeURIComponent('oracleid:' + oracleId + ' lang:en');
    } else {
      url =
        'https://api.scryfall.com/cards/named?fuzzy=' +
        encodeURIComponent(cleanCardName(cardName));
    }

    const response =
      UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          'User-Agent':
            'MTG-Card-Scanner/1.0 contact: mtg-card-scanner',
          'Accept':
            'application/json;q=0.9,*/*;q=0.8'
        },
        muteHttpExceptions: true
      });

    if (response.getResponseCode() === 200) {
      const data =
        JSON.parse(response.getContentText());
      if (data.data && data.data.length > 0) {
        return cleanCardName(data.data[0].name || '');
      }
      if (data.name) {
        return cleanCardName(data.name);
      }
    }
  } catch (e) {
    console.warn('英語名補完失敗: ' + e.message);
  }
  return '';
}

/**
 * Scryfallカードオブジェクトから日本語名・英語名を抽出・補完
 * ・元から判定できているものはそのまま使い、埋まっていないものだけ補完を行う
 * ・「//」が含まれる場合はそれ以前の文字列のみを格納する
 */
function extractCardNames(card) {
  let englishName = '';
  let japaneseName = '';

  // 1. 英語オラクル名
  if (card.name) {
    englishName = cleanCardName(card.name);
  } else if (card.card_faces && card.card_faces[0] && card.card_faces[0].name) {
    englishName = cleanCardName(card.card_faces[0].name);
  }

  // 2. 渡されたカードオブジェクト自体が日本語版の場合
  if (card.lang === 'ja' || card.lang === 'japanese') {
    if (card.printed_name) {
      japaneseName = cleanCardName(card.printed_name);
    } else if (card.card_faces && card.card_faces[0] && card.card_faces[0].printed_name) {
      japaneseName = cleanCardName(card.card_faces[0].printed_name);
    }
  }

  // 3. 日本語名が埋まっていない場合のみ、Scryfall検索で補完
  if (!japaneseName) {
    const fetchedJa =
      fetchJapaneseName(card.oracle_id, englishName);
    if (fetchedJa) {
      japaneseName = cleanCardName(fetchedJa);
    }
  }

  // 4. 英語名が埋まっていない場合のみ、Scryfall検索で補完
  if (!englishName) {
    const fetchedEn =
      fetchEnglishName(card.oracle_id, japaneseName);
    if (fetchedEn) {
      englishName = cleanCardName(fetchedEn);
    }
  }

  // 5. 日本語版印刷が存在しないカード（英語限定カード等）のフォールバック
  if (!japaneseName) {
    japaneseName = englishName;
  }
  if (!englishName) {
    englishName = japaneseName;
  }

  return {
    card_name: cleanCardName(japaneseName),
    card_english_name: cleanCardName(englishName)
  };
}

/**
 * Scryfallのカードデータを
 * アプリ内部で使う形式に変換
 */
function convertScryfallCard(card) {

  if (!card) {
    throw new Error(
      'Scryfallカード情報がありません'
    );
  }

  const color =
    getCardColor(card);

  const type =
    getCardType(card);

  let manaValue = '';

  if (card.mana_value != null) {
    manaValue =
      card.mana_value;
  } else if (card.cmc != null) {
    manaValue =
      card.cmc;
  }

  const names =
    extractCardNames(card);

  const result = {
    card_name:
      names.card_name,

    card_english_name:
      names.card_english_name,

    set_code:
      card.set || '',

    collector_number:
      card.collector_number || '',

    language:
      card.lang || '',

    foil:
      false,

    color:
      color,

    type:
      type,

    mana_value:
      manaValue
  };

  console.log(
    '===== Scryfall変換結果 ====='
  );

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}

/**
 * Gemini APIを呼び出す
 *
 * メインモデルで429が発生した場合、
 * フォールバックモデルへ切り替える。
 *
 * @param {Object} payload Gemini API payload
 * @return {Object} Gemini API response JSON
 */
function callGeminiWithFallback(payload) {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error(
      "GEMINI_API_KEY が設定されていません"
    );
  }


  // 通常使用するモデル
  const primaryModel =
    "gemini-3.5-flash-lite";

  // 429時のフォールバック
  const fallbackModel =
    "gemini-3.1-flash-lite";


  /*
   * まずメインモデル
   */
  let response =
    callGeminiModel(
      primaryModel,
      payload,
      apiKey
    );


  /*
   * 成功
   */
  if (response.status >= 200 &&
      response.status < 300) {

    console.log(
      "Gemini使用モデル:",
      primaryModel
    );

    return JSON.parse(
      response.body
    );
  }


  /*
   * 429/503ならフォールバック
   */
  if (response.status === 429 || response.status === 503) {

    console.warn(
      primaryModel +
      " が429を返しました。"
    );

    console.warn(
      "フォールバックモデル " +
      fallbackModel +
      " を使用します。"
    );


    response =
      callGeminiModel(
        fallbackModel,
        payload,
        apiKey
      );


    /*
     * フォールバック成功
     */
    if (
      response.status >= 200 &&
      response.status < 300
    ) {

      console.log(
        "Gemini使用モデル:",
        fallbackModel
      );

      return JSON.parse(
        response.body
      );
    }


    /*
     * フォールバックも失敗
     */
    throw new Error(
      "Gemini APIエラー\n" +
      "Primary: " +
      primaryModel +
      " → HTTP " +
      "429\n" +
      "Fallback: " +
      fallbackModel +
      " → HTTP " +
      response.status +
      "\n" +
      response.body
    );
  }


  /*
   * 429以外
   */
  throw new Error(
    "Gemini APIエラー\n" +
    "Model: " +
    primaryModel +
    "\n" +
    "HTTP Status: " +
    response.status +
    "\n" +
    response.body
  );
}


/**
 * 指定したGeminiモデルを呼び出す
 */
function callGeminiModel(
  model,
  payload,
  apiKey
) {

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent";


  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: "post",

        contentType:
          "application/json",

        headers: {
          "x-goog-api-key":
            apiKey
        },

        payload:
          JSON.stringify(payload),

        muteHttpExceptions:
          true
      }
    );


  return {
    status:
      response.getResponseCode(),

    body:
      response.getContentText()
  };
}
