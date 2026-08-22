/**
 * カード検索関数
 *
 * @param {Object} criteria 検索条件
 *   - query: string (カード名・日本語名または英語名の部分一致)
 *   - colors: Array<string> (色: '白', '青', '黒', '赤', '緑', '無色')
 *   - manaValue: string|number (マナコスト: 0〜6, '7+')
 *   - minPriceMin: number|string (最低価格 下限)
 *   - minPriceMax: number|string (最低価格 上限)
 *   - status: string (ステータス: 'deck', 'case', 'storage', 'selling')
 *   - deck_id: string (デッキID)
 * @return {Array<Object>} 検索結果のカード一覧
 */
function searchCards(criteria) {
  criteria = criteria || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CARDS);

  if (!sheet) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const columns = getColumnMap(sheet);
  const data = sheet.getDataRange().getValues();

  // DeckSummary のマップを取得（deck_name 表示用）
  const deckMap = {};
  try {
    const decks = typeof getDecks === 'function' ? getDecks() : [];
    decks.forEach(function(d) {
      if (d.deck_id) {
        deckMap[d.deck_id] = d.deck_name || d.deck_id;
      }
    });
  } catch (e) {
    console.warn('デッキマップ取得エラー:', e.message);
  }

  // 検索条件のパース
  const query = String(criteria.query || '').trim().toLowerCase();
  const selectedColors = Array.isArray(criteria.colors)
    ? criteria.colors.filter(Boolean).map(function(c) { return String(c).trim(); })
    : [];
  const targetMana = criteria.manaValue !== undefined && criteria.manaValue !== null && criteria.manaValue !== ''
    ? String(criteria.manaValue).trim()
    : '';

  const minPriceMin = criteria.minPriceMin !== undefined && criteria.minPriceMin !== null && criteria.minPriceMin !== ''
    ? Number(criteria.minPriceMin)
    : null;
  const minPriceMax = criteria.minPriceMax !== undefined && criteria.minPriceMax !== null && criteria.minPriceMax !== ''
    ? Number(criteria.minPriceMax)
    : null;

  const targetStatus = criteria.status && criteria.status !== 'all'
    ? String(criteria.status).trim().toLowerCase()
    : '';
  const targetDeckId = criteria.deck_id
    ? String(criteria.deck_id).trim()
    : '';

  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const cardName = columns['card_name'] ? String(row[columns['card_name'] - 1] || '').trim() : '';
    const cardEnglishName = columns['card_english_name'] ? String(row[columns['card_english_name'] - 1] || '').trim() : '';
    const color = columns['color'] ? String(row[columns['color'] - 1] || '').trim() : '';
    const type = columns['type'] ? String(row[columns['type'] - 1] || '').trim() : '';
    const manaValueRaw = columns['mana_value'] ? row[columns['mana_value'] - 1] : '';
    const setCode = columns['set_code'] ? String(row[columns['set_code'] - 1] || '').trim() : '';
    const collectorNumber = columns['collector_number'] ? String(row[columns['collector_number'] - 1] || '').trim() : '';
    const language = columns['language'] ? String(row[columns['language'] - 1] || '').trim() : '';
    const foil = columns['foil'] ? Boolean(row[columns['foil'] - 1]) : false;
    const status = columns['status'] ? String(row[columns['status'] - 1] || '').trim().toLowerCase() : '';
    const deckId = columns['deck_id'] ? String(row[columns['deck_id'] - 1] || '').trim() : '';
    const count = columns['count'] ? Number(row[columns['count'] - 1]) || 1 : 1;
    const minPriceRaw = columns['min_price'] ? row[columns['min_price'] - 1] : '';
    const avgPriceRaw = columns['avg_price'] ? row[columns['avg_price'] - 1] : '';
    const updatedAt = columns['updated_at'] ? row[columns['updated_at'] - 1] : '';

    const minPrice = (minPriceRaw !== '' && minPriceRaw !== null && !isNaN(minPriceRaw))
      ? Number(minPriceRaw)
      : null;
    const avgPrice = (avgPriceRaw !== '' && avgPriceRaw !== null && !isNaN(avgPriceRaw))
      ? Number(avgPriceRaw)
      : null;

    // 1. カード名フィルター (部分一致)
    if (query) {
      const matchName = cardName.toLowerCase().indexOf(query) !== -1;
      const matchEnglish = cardEnglishName.toLowerCase().indexOf(query) !== -1;
      if (!matchName && !matchEnglish) {
        continue;
      }
    }

    // 2. 色フィルター
    if (selectedColors.length > 0) {
      // カードの色（例: '白,青', '無色'）
      const cardColors = color ? color.split(',').map(function(c) { return c.trim(); }) : [];
      let colorMatched = false;

      // 選択された色のいずれかを含んでいるか判定
      for (let c = 0; c < selectedColors.length; c++) {
        const sc = selectedColors[c];
        if (sc === '無色') {
          if (cardColors.indexOf('無色') !== -1 || cardColors.length === 0) {
            colorMatched = true;
            break;
          }
        } else if (cardColors.indexOf(sc) !== -1) {
          colorMatched = true;
          break;
        }
      }

      if (!colorMatched) {
        continue;
      }
    }

    // 3. マナコストフィルター
    if (targetMana !== '') {
      const cardManaNum = (manaValueRaw !== '' && manaValueRaw !== null && !isNaN(manaValueRaw))
        ? Number(manaValueRaw)
        : null;

      if (targetMana === '7+') {
        if (cardManaNum === null || cardManaNum < 7) {
          continue;
        }
      } else {
        const targetManaNum = Number(targetMana);
        if (cardManaNum !== targetManaNum) {
          continue;
        }
      }
    }

    // 4. 最低価格 (min_price) レンジフィルター
    if (minPriceMin !== null) {
      if (minPrice === null || minPrice < minPriceMin) {
        continue;
      }
    }
    if (minPriceMax !== null) {
      if (minPrice === null || minPrice > minPriceMax) {
        continue;
      }
    }

    // 5. ステータスフィルター
    if (targetStatus && status !== targetStatus) {
      continue;
    }

    // 6. デッキIDフィルター
    if (targetDeckId && deckId !== targetDeckId) {
      continue;
    }

    results.push({
      card_name: cardName,
      card_english_name: cardEnglishName,
      color: color,
      type: type,
      mana_value: manaValueRaw,
      set_code: setCode,
      collector_number: collectorNumber,
      language: language,
      foil: foil,
      status: status,
      deck_id: deckId,
      deck_name: deckMap[deckId] || '',
      count: count,
      min_price: minPrice,
      avg_price: avgPrice,
      updated_at: updatedAt ? Utilities.formatDate(new Date(updatedAt), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : ''
    });

    // 最大200件で打ち切り
    if (results.length >= 200) {
      break;
    }
  }

  return results;
}
