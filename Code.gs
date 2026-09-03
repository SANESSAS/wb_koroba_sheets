function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Заполнить ID по коробам', 'zapolnitShkKorobov')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function zapolnitShkKorobov() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ищем лист ID, игнорируя случайные пробелы и регистр
  const idSheet = ss.getSheets().find(
    s => s.getName().trim().toUpperCase() === 'ID'
  );

  const dataSheet = ss.getSheets().find(
    s => s.getName().trim().toLowerCase() === 'готовые данные'
  );

  // Если ID не найден — сразу покажет,
  // в какой таблице запущен скрипт и какие листы там есть
  if (!idSheet) {
    throw new Error(
      'Не найден лист "ID".\n\n' +
      'Текущая таблица: ' + ss.getName() + '\n\n' +
      'Найденные листы:\n' +
      ss.getSheets().map(s => '[' + s.getName() + ']').join('\n')
    );
  }

  if (!dataSheet) {
    throw new Error(
      'Не найден лист "Готовые данные".\n\n' +
      'Текущая таблица: ' + ss.getName()
    );
  }

  // ==========================================
  // 1. Читаем ID из колонки A сверху вниз
  // ==========================================
  const lastRow = idSheet.getLastRow();

  const ids = idSheet
    .getRange(1, 1, lastRow, 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v).trim())
    .filter(v => v !== '');

  if (ids.length === 0) {
    throw new Error('На листе ID колонка A пустая.');
  }

  // ==========================================
  // 2. Ищем короба в строке 4
  // ВАЖНО: НЕ сортируем!
  // Порядок колонок слева направо сохраняется.
  // ==========================================
  const lastColumn = dataSheet.getLastColumn();

  const headers = dataSheet
    .getRange(4, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const boxColumns = [];

  for (let col = 0; col < headers.length; col++) {
    const text = String(headers[col]).trim();

    // Короб 1, Короб 2, Короб 3 и т.д.
    if (/^Короб\s*\d+$/i.test(text)) {
      boxColumns.push(col + 1); // настоящий номер колонки Sheets
    }
  }

  if (boxColumns.length === 0) {
    throw new Error(
      'В строке 4 листа "Готовые данные" не найдено ни одного "Короб N".'
    );
  }

  if (ids.length > boxColumns.length) {
    throw new Error(
      'ID больше, чем найденных коробов.\n\n' +
      'ID: ' + ids.length + '\n' +
      'Коробов: ' + boxColumns.length
    );
  }

  // ==========================================
  // 3. Записываем ID в строку 2
  //
  // ID №1 → первый физический Короб слева
  // ID №2 → второй физический Короб слева
  // ID №3 → третий и т.д.
  //
  // Используем только нужные ячейки,
  // остальные данные листа не перезаписываем.
  // ==========================================
  for (let i = 0; i < ids.length; i++) {
    dataSheet
      .getRange(2, boxColumns[i])
      .setValue(ids[i]);
  }

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Готово ✅\n\n' +
    'Таблица: ' + ss.getName() + '\n' +
    'ID: ' + ids.length + '\n' +
    'Найдено позиций коробов: ' + boxColumns.length + '\n\n' +
    'Заполнены первые ' + ids.length + ' коробов слева направо.'
  );
}
