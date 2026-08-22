/**
 * 既存のCardsシートを1行ずつ確認し、
 * card_name（日本語名）または card_english_name（英語名）を補完するバッチ関数
 *
 * 既存DBのデータ更新のため手動で1度実行して利用します。
 */
function migrateCardNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CARDS);

  if (!sheet) {
    throw new Error('Cardsシートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('Cardsシートにデータが存在しません');
    return { success: true, message: 'データなし', total: 0, updated: 0 };
  }

  // カラムマップを取得
  let columns = getColumnMap(sheet);

  // card_english_name カラムが存在しない場合は自動追加
  if (!columns['card_english_name']) {
    const lastCol = sheet.getLastColumn();
    sheet.getRange(1, lastCol + 1).setValue('card_english_name');
    columns = getColumnMap(sheet);
  }

  const cardNameCol = columns['card_name'];
  const englishNameCol = columns['card_english_name'];
  const setCodeCol = columns['set_code'];
  const collectorNumberCol = columns['collector_number'];
  const colorCol = columns['color'];
  const typeCol = columns['type'];
  const manaValueCol = columns['mana_value'];

  if (!cardNameCol || !englishNameCol) {
    throw new Error('card_name または card_english_name 列が存在しません');
  }

  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  Logger.log('=== Cardsシート名称補完バッチ開始 (全 ' + (data.length - 1) + ' 件) ===');

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const currentName = String(data[i][cardNameCol - 1] || '').trim();
    const currentEnglishName = String(data[i][englishNameCol - 1] || '').trim();
    const setCode = setCodeCol ? String(data[i][setCodeCol - 1] || '').trim() : '';
    const collectorNumber = collectorNumberCol ? String(data[i][collectorNumberCol - 1] || '').trim() : '';

    // すでに日本語名と英語名の両方が異なる値で入っている場合はスキップ
    // （ただし英語名が空、または両方が同一の場合は補完対象）
    const needsUpdate = !currentName || !currentEnglishName || currentName === currentEnglishName;

    if (!needsUpdate) {
      skippedCount++;
      continue;
    }

    const searchKey = currentName || currentEnglishName;
    if (!searchKey && (!setCode || !collectorNumber)) {
      skippedCount++;
      continue;
    }

    try {
      let scryfallCard = null;

      // 1. セットコードとコレクター番号で完全照合
      if (setCode && collectorNumber) {
        try {
          scryfallCard = searchScryfallBySetAndCollector(setCode, collectorNumber);
        } catch (e) {
          Logger.log('Row ' + row + ': セット照合失敗 (' + setCode + '/' + collectorNumber + ') -> 名前検索へフォールバック');
        }
      }

      // 2. 完全照合ができなかった場合はカード名で検索
      if (!scryfallCard && searchKey) {
        scryfallCard = searchScryfallCard(searchKey);
      }

      if (scryfallCard) {
        const jaName = scryfallCard.card_name || currentName || '';
        const enName = scryfallCard.card_english_name || currentEnglishName || '';

        // シートを更新
        if (jaName) {
          sheet.getRange(row, cardNameCol).setValue(jaName);
        }
        if (enName) {
          sheet.getRange(row, englishNameCol).setValue(enName);
        }

        // color, type, mana_value も空なら補完
        if (colorCol && !data[i][colorCol - 1] && scryfallCard.color) {
          sheet.getRange(row, colorCol).setValue(scryfallCard.color);
        }
        if (typeCol && !data[i][typeCol - 1] && scryfallCard.type) {
          sheet.getRange(row, typeCol).setValue(scryfallCard.type);
        }
        if (manaValueCol && (data[i][manaValueCol - 1] === '' || data[i][manaValueCol - 1] === null) && scryfallCard.mana_value !== '') {
          sheet.getRange(row, manaValueCol).setValue(scryfallCard.mana_value);
        }

        updatedCount++;
        Logger.log('Row ' + row + ' 更新成功: 日本語名=[' + jaName + '] / 英語名=[' + enName + ']');
      } else {
        Logger.log('Row ' + row + ' スキップ (Scryfallで情報が取得できませんでした): ' + searchKey);
        errorCount++;
      }

      // Scryfall API レートリミット対策
      Utilities.sleep(100);

    } catch (err) {
      Logger.log('Row ' + row + ' エラー: ' + err.message);
      errorCount++;
    }
  }

  const resultMessage = '=== 補完バッチ完了 ===\n' +
    '全データ行: ' + (data.length - 1) + ' 行\n' +
    '更新件数: ' + updatedCount + ' 件\n' +
    'スキップ: ' + skippedCount + ' 件\n' +
    'エラー/未取得: ' + errorCount + ' 件';

  Logger.log(resultMessage);

  return {
    success: true,
    message: resultMessage,
    total: data.length - 1,
    updated: updatedCount,
    skipped: skippedCount,
    errors: errorCount
  };
}
