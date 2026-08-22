function fetchPrice(cardName) {
  const url =
    "https://whisper.wisdom-guild.net/card/" +
    encodeURIComponent(cardName);

  const response =
    UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0"
      }
    });

  const status =
    response.getResponseCode();

  if (status !== 200) {
    throw new Error(
      "Wisdom Guild取得エラー: HTTP " +
      status
    );
  }

  const content =
    response.getContentText("utf-8");

  /*
   * 最安
   */
  let minPriceStr =
    Parser
      .data(content)
      .from('<big>最安：<b>')
      .to('</b>')
      .iterate();

  minPriceStr =
    String(minPriceStr)
      .replace(/,/g, '')
      .replace(/円/g, '')
      .trim();

  const minPrice =
    parseInt(minPriceStr, 10);

  /*
   * トリム平均
   */
  let avgPriceStr =
    Parser
      .data(content)
      .from('／トリム平均：<b>')
      .to('</b>')
      .iterate();

  avgPriceStr =
    String(avgPriceStr)
      .replace(/,/g, '')
      .replace(/円/g, '')
      .trim();

  const avgPrice =
    parseInt(avgPriceStr, 10);

  if (
    isNaN(minPrice) ||
    isNaN(avgPrice)
  ) {
    throw new Error(
      "価格を取得できませんでした: " +
      cardName
    );
  }

  return {
    minPrice: minPrice,
    avgPrice: avgPrice
  };
}


/**
 * Cardsシートのヘッダーから列番号を取得
 *
 * 例：
 * scryfall_id → 1
 * card_name   → 2
 * color       → 3
 * ...
 */
function getPriceColumnMap(sheet) {

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];

  const columns = {};

  headers.forEach(function(header, index) {

    if (header) {

      columns[
        String(header).trim()
      ] = index + 1;

    }

  });

  return columns;
}


function updateAllCardPrices() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(SHEET_CARDS);

  const logSheet =
    ss.getSheetByName(SHEET_PRICE_LOG);


  if (!sheet) {
    throw new Error(
      'Cardsシートがありません'
    );
  }


  if (!logSheet) {
    throw new Error(
      'PriceLogシートがありません。' +
      'setupPriceLog()を先に実行してください'
    );
  }


  /*
   * Cardsシートのヘッダーから
   * 各列の位置を取得
   */
  const columns =
    getPriceColumnMap(sheet);


  /*
   * 必須カラム
   */
  const requiredColumns = [
    'card_name',
    'set_code',
    'collector_number',
    'min_price',
    'avg_price',
    'previous_min_price',
    'previous_avg_price',
    'price_updated_at'
  ];


  /*
   * 必須カラムが存在するか確認
   */
  requiredColumns.forEach(function(column) {

    if (!columns[column]) {

      throw new Error(
        'Cardsシートに「' +
        column +
        '」列がありません'
      );

    }

  });


  /*
   * Cardsシートのデータ取得
   */
  const data =
    sheet
      .getDataRange()
      .getValues();


  const notifications = [];


  /*
   * 2行目から処理
   */
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      i + 1;


    /*
     * カード名をヘッダー名から取得
     */
    const cardName =
      data[i][
        columns['card_name'] - 1
      ];


    if (!cardName) {
      continue;
    }


    try {

      Logger.log(
        "価格取得: " +
        cardName
      );


      /*
       * 現在価格を取得
       */
      const price =
        fetchPrice(cardName);


      /*
       * 前回の最安価格
       */
      const previousMin =
        Number(
          data[i][
            columns['min_price'] - 1
          ]
        ) || 0;


      /*
       * 前回のトリム平均
       */
      const previousAvg =
        Number(
          data[i][
            columns['avg_price'] - 1
          ]
        ) || 0;


      /*
       * 前回価格を保存
       */
      sheet
        .getRange(
          row,
          columns['previous_min_price']
        )
        .setValue(
          previousMin
        );


      sheet
        .getRange(
          row,
          columns['previous_avg_price']
        )
        .setValue(
          previousAvg
        );


      /*
       * 現在の最安価格
       */
      sheet
        .getRange(
          row,
          columns['min_price']
        )
        .setValue(
          price.minPrice
        );


      /*
       * 現在のトリム平均
       */
      sheet
        .getRange(
          row,
          columns['avg_price']
        )
        .setValue(
          price.avgPrice
        );


      /*
       * 価格更新日時
       */
      sheet
        .getRange(
          row,
          columns['price_updated_at']
        )
        .setValue(
          new Date()
        );


      /*
       * PriceLog
       *
       * PriceLog側は従来通りの固定順
       */
      logSheet.appendRow([

        new Date(),

        cardName,

        columns['card_english_name']
          ? data[i][columns['card_english_name'] - 1]
          : '',

        data[i][
          columns['set_code'] - 1
        ],

        data[i][
          columns['collector_number'] - 1
        ],

        price.minPrice,

        price.avgPrice

      ]);


      /*
       * 値上がり判定
       *
       * 最安価格が
       * 300円以上になり、
       * 前回より上がった場合
       */
      if (
        previousMin > 0 &&
        price.minPrice > previousMin &&
        price.minPrice >= 300
      ) {

        notifications.push({

          cardName:
            cardName,

          beforePrice:
            previousMin,

          currentPrice:
            price.minPrice,

          difference:
            price.minPrice -
            previousMin

        });

      }


    } catch (error) {

      Logger.log(
        "価格取得失敗: " +
        cardName +
        " / " +
        error.message
      );

    }


    /*
     * Wisdom Guildへのアクセス間隔
     */
    Utilities.sleep(1000);

  }


  /*
   * メール通知
   */
  if (
    notifications.length > 0
  ) {

    sendPriceIncreaseEmail(
      notifications
    );

  }


  return {

    success: true,

    updated:
      data.length - 1,

    notifications:
      notifications.length

  };

}


function sendPriceIncreaseEmail(
  notifications
) {

  const to =
    "paperbackwriter0049@gmail.com";


  let body =
    "MTGカードの価格が上昇しました。\n\n";


  for (
    const item of notifications
  ) {

    body +=
      "カード名: " +
      item.cardName +
      "\n" +

      "前回: " +
      item.beforePrice +
      "円\n" +

      "現在: " +
      item.currentPrice +
      "円\n" +

      "上昇額: +" +
      item.difference +
      "円\n\n";

  }


  GmailApp.sendEmail(
    to,
    "【MTG】カード価格上昇通知",
    body
  );

}