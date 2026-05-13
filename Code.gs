function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return ContentService.createTextOutput(JSON.stringify(runAction(e.parameter.action, e.parameter)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Бронирование Mazaysplav')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function apiHandler(action, params) {
  try {
    return runAction(action, params);
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function runAction(action, params) {
  let result;
  switch (action) {
    case 'getDates':
      result = getAvailableDates();
      break;
    case 'checkExistingBooking':
      result = checkExistingBooking(params.phone, params.tripName || "");
      break;
    case 'processBooking':
      let data = params.data;
      if (typeof data === 'string') data = JSON.parse(decodeURIComponent(data));
      result = processBooking(data);
      break;
    case 'cancelBooking':
      result = cancelBooking(params.phone, params.tripName);
      break;
    default:
      result = { error: 'Unknown action: ' + action };
  }
  return result;
}

// ============================================
// НАСТРОЙКИ
// ============================================
const CONFIG = {
  BOT_TOKEN: '8759405862:AAHeVKwNMJD1k-Ehnyx7PORxEiH-Y7aSM50',
  ADMIN_CHAT_ID: '336999745',
  FLEET_2M: 40,
  FLEET_3M: 6,
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
};

function getAvailableDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const splav = ss.getSheetByName('Сплавы');
    const lastRow = splav.getLastRow();
    const lastCol = splav.getLastColumn();
    const data = splav.getRange(1, 1, lastRow, Math.max(lastCol, 52)).getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    const getVal = (row, name) => {
      const idx = headers.indexOf(name.toLowerCase());
      return idx >= 0 ? row[idx] : "";
    };

    const occByDate = data.slice(1).reduce((acc, row) => {
      const d = row[headers.indexOf('дата')];
      if (d) {
        const dStr = Utilities.formatDate(new Date(d), tz, "dd.MM.yyyy");
        if (!acc[dStr]) acc[dStr] = { c: 0, t: 0, s: 0, b: 0, p: 0 };
        acc[dStr].c += (parseInt(getVal(row, 'Занято Кемперов')) || 0);
        acc[dStr].t += (parseInt(getVal(row, 'Занято Типи')) || 0);
        acc[dStr].s += (parseInt(getVal(row, 'Занято Сафари')) || 0);
        acc[dStr].b += (parseInt(getVal(row, 'Занято Бань')) || 0);
        
        let tentColIdx = headers.indexOf('занято палаток (ваших)');
        if (tentColIdx < 0) tentColIdx = headers.indexOf('занято палаток (наших)');
        if (tentColIdx >= 0) acc[dStr].p += (parseInt(row[tentColIdx]) || 0);
      }
      return acc;
    }, {});
    
    const dailyOccupied = {}; 
    const tripDetails = [];   

    for (let i = 1; i < data.length; i++) {
      const rawDate = getVal(data[i], 'Дата');
      if (!rawDate) { tripDetails.push({b2:0, b3:0, tr:0, dateStr:""}); continue; }
      
      const dateStr = Utilities.formatDate(new Date(rawDate), tz, "dd.MM.yyyy");
      const route = String(getVal(data[i], 'Маршрут') || 'Сплав').trim();
      let sName = dateStr;
      const tVal = getVal(data[i], 'Время');
      let timeStr = "";
      if (tVal) {
        if (tVal instanceof Date) timeStr = Utilities.formatDate(tVal, tz, "HH:mm");
        else {
          const tParts = String(tVal).split(':');
          if (tParts.length >= 2) timeStr = tParts[0].trim().padStart(2, '0') + ":" + tParts[1].trim().substring(0, 2);
          else timeStr = String(tVal).split(' ')[0];
        }
        sName += " " + timeStr;
      }
      // Добавляем реку к имени листа для поиска
      // Формируем имя без удаления двоеточия
      sName = (sName + " " + route).trim().replace(/[\\\/\?\*\[\]]/g, '').substring(0, 31);
      
      // Тщательный поиск листа
      let tripSheet = ss.getSheetByName(sName);
      if (!tripSheet) {
        const allSheets = ss.getSheets();
        // Сравнение с удалением всех пробелов и двоеточий для 100% попадания
        const clean = (s) => s.replace(/[\s:]/g, '').toLowerCase();
        tripSheet = allSheets.find(s => clean(s.getName()) === clean(sName));
      }

      let b2 = 0, b3 = 0, tr = 0;
      if (tripSheet) {
        const tData = tripSheet.getDataRange().getValues();
        for (let j = 1; j < tData.length; j++) {
          b2 += (parseFloat(tData[j][6]) || 0); 
          b3 += (parseFloat(tData[j][7]) || 0); 
          tr += (parseFloat(tData[j][8]) || 0);
        }
        logAction('DEBUG_OCCUPANCY', `Found: ${tripSheet.getName()}, B2: ${b2}`, 'Success');
      } else {
        logAction('DEBUG_OCCUPANCY', `NOT FOUND: [${sName}]`, 'Warning');
      }
      
      tripDetails.push({b2, b3, tr, dateStr, sName, timeStr});
      if (!dailyOccupied[dateStr]) dailyOccupied[dateStr] = { b2: 0, b3: 0, tr: 0 };
      dailyOccupied[dateStr].b2 += b2;
      dailyOccupied[dateStr].b3 += b3;
      dailyOccupied[dateStr].tr += tr;
    }

    const results = [];
    const blueValues = [];

    for (let i = 1; i < data.length; i++) {
      const details = tripDetails[i-1];
      const dateStr = details ? details.dateStr : "";
      if (!dateStr) { blueValues.push(["","","",""]); continue; }

      const rawLimP = getVal(data[i], 'Лимит людей');
      const rawLim2 = getVal(data[i], 'Лимит байдарка 2М');
      const rawLim3 = getVal(data[i], 'Лимит байдарка 3М');
      const rawLimT = getVal(data[i], 'Лимит трансфер');

      const limP = (rawLimP === "") ? -1 : parseInt(rawLimP);
      const lim2 = (rawLim2 === "") ? -1 : parseInt(rawLim2);
      const lim3 = (rawLim3 === "") ? -1 : parseInt(rawLim3);
      const limT = (rawLimT === "") ? -1 : parseInt(rawLimT);

      const dayOcc = dailyOccupied[dateStr]; 

      let remP  = limP >= 0 ? Math.max(0, limP - (details.b2*2 + details.b3*3)) : "";
      let rem2M = lim2 >= 0 ? Math.max(0, lim2 - details.b2) : Math.max(0, CONFIG.FLEET_2M - dayOcc.b2);
      let rem3M = lim3 >= 0 ? Math.max(0, lim3 - details.b3) : Math.max(0, CONFIG.FLEET_3M - dayOcc.b3);
      let remTr = limT >= 0 ? Math.max(0, limT - details.tr) : 999;

      blueValues.push([remP, rem2M, rem3M, remTr === 999 ? "∞" : remTr]);

      if (String(getVal(data[i], 'Активен')).toLowerCase() === 'да') {
        results.push({
          date: dateStr, time: details.timeStr,
          dayType: String(getVal(data[i], 'Дни') || "").toLowerCase(), 
          route: String(getVal(data[i], 'Маршрут') || 'Сплав'),
          type: String(getVal(data[i], 'Тип') || '1 день'),
          gps: String(getVal(data[i], 'Координаты') || getVal(data[i], 'GPS') || ''),
          hasFood: String(getVal(data[i], 'Пропитание') || getVal(data[i], 'Питание') || '').toLowerCase().includes('да'),
          rem2M: rem2M, rem3M: rem3M, remTransfer: remTr,
          priceAdultNo: parseFloat(getVal(data[i], 'Взр. без трансфера')) || 0,
          priceAdultWith: parseFloat(getVal(data[i], 'Взр. с трансфером')) || 0,
          priceChildNo: parseFloat(getVal(data[i], 'Реб. без трансфера')) || 0,
          priceChildWith: parseFloat(getVal(data[i], 'Реб. с трансфером')) || 0,
          program: String(getVal(data[i], 'Программа') || ''),
          pamyatka: String(getVal(data[i], 'Памятка') || ''),
          occCamper: (occByDate[dateStr] ? occByDate[dateStr].c : 0),
          occTipi: (occByDate[dateStr] ? occByDate[dateStr].t : 0),
          occSafari: (occByDate[dateStr] ? occByDate[dateStr].s : 0),
          occSauna: (occByDate[dateStr] ? occByDate[dateStr].b : 0),
          occOurTent: (occByDate[dateStr] ? occByDate[dateStr].p : 0),
          priceCamper: parseFloat(getVal(data[i], 'Цена Кемпера')) || 0,
          priceTipi: parseFloat(getVal(data[i], 'Цена Типи')) || 0,
          priceSafari: parseFloat(getVal(data[i], 'Цена Сафари')) || 0,
          priceOurTent: parseFloat(getVal(data[i], 'Цена Вашей Палатки')) || 0,
          priceOwnTent: parseFloat(getVal(data[i], 'Цена Своей Палатки')) || 0,
          priceSaunaHot: parseFloat(getVal(data[i], 'Цена Баня+Купель')) || 0,
        });
      }
    }
    try {
      if (blueValues.length > 0) splav.getRange(2, 19, blueValues.length, 4).setValues(blueValues);
    } catch (err) {
      console.error("Ошибка при обновлении остатков в таблице: " + err);
    }

    return results;
  } catch (e) { 
    return { error: e.toString() }; 
  }
}

function processBooking(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const adults = parseInt(data.adults) || 0;
    const children = parseInt(data.children) || 0;
    const total = adults + children;
    const priceA = parseFloat(data.unitPriceAdult) || 0;
    const priceC = parseFloat(data.unitPriceChild) || 0;
    const totalSum = parseFloat(data.sum) || 0;
    const cleanPhone = data.phone.replace(/\D/g,'');
    const bookingId = "M-" + Math.random().toString(36).substr(2, 6).toUpperCase();

    const rawRoute = String(data.route || "Сплав").trim();
    const cleanRoute = rawRoute.replace(/[\\\/\?\*\[\]]/g, '');
    let sheetName = (data.date + (data.time ? " " + data.time : "") + " " + cleanRoute).trim().substring(0, 31);

    logAction('DEBUG_BOOKING', `Route: ${rawRoute}, Date: ${data.date}, Time: ${data.time}, Sheet: ${sheetName}`, 'Info');

    // Ищем лист именно для этого маршрута
    let tripSheet = ss.getSheetByName(sheetName);
    
    if (!tripSheet) {
      tripSheet = ss.insertSheet(sheetName);
      const headerFormulas = [
        'Имя', 'Телефон', 'Маршрут', 'Тип', 
        '="Взрослых: "&SUM(E2:E1000)', '="Детей: "&SUM(F2:F1000)', '="Байд2М: "&SUM(G2:G1000)', '="Байд3М: "&SUM(H2:H1000)', '="Трансфер: "&SUM(I2:I1000)', 
        'Собака', 'Комментарий', 'Жилье', 'Доп. услуги', 'Цена взр.', 'Цена дет.', 
        '="ИТОГО: "&SUM(P2:P1000)', 'ID Брони', 'Оплата доставки', 'Остаток', 'Статус'
      ];
      tripSheet.appendRow(headerFormulas);
      tripSheet.getRange('A1:T1').setFontWeight('bold').setBackground('#dcfce7').setBorder(true, true, true, true, true, true);
      tripSheet.setFrozenRows(1);
    }
    const isWaitlist = data.isWaitlist === true;
    const prepayment = (!isWaitlist && data.dayType && data.dayType.includes('вых') && data.transfer !== 'Да') ? (total * 20) : 0;
    const remaining = totalSum - prepayment;

    const formatV = (val) => (isWaitlist && val > 0) ? "v" + String(val).replace('.', ',') : val;

    const adultsCell = formatV(adults);
    const childrenCell = formatV(children);
    const k2Cell = formatV(parseFloat(data.k2) || 0);
    const k3Cell = formatV(parseFloat(data.k3) || 0);
    const transferCount = data.transfer === 'Да' ? total : 0;
    const transferCell = formatV(transferCount);
    
    const priceACell = formatV(adults > 0 ? priceA : 0);
    const priceCCell = formatV(children > 0 ? priceC : 0);
    const totalSumCell = formatV(totalSum);
    const remainingCell = formatV(remaining);
    const prepaymentCell = formatV(prepayment);
    const foodStr = data.foodPrefs ? ` | Питание: ${data.foodPrefs}` : "";
    const finalComment = (data.comment + foodStr).trim();

    const lastRow = tripSheet.getLastRow() + 1;
    const rowData = [
      data.name, "'" + data.phone, data.route, data.type, adultsCell, childrenCell, k2Cell, k3Cell, transferCell, 
      data.dog || "нет", finalComment, 
      data.accommodation || "-", data.extraServices || "-", 
      priceACell, priceCCell, totalSumCell, bookingId, prepaymentCell, remainingCell, "Ожидает"
    ];
    tripSheet.appendRow(rowData);
    
    // ПРОВЕРКА НА МНОЖЕСТВЕННОЕ БРОНИРОВАНИЕ
    const checkMulti = checkExistingBooking(data.phone, "");
    if (checkMulti.found && checkMulti.allCount > 1) {
       repaintMultiBookings(data.phone);
    } else {
       // Если бронь одна - красим по стандартным правилам
       if (data.hasTelegram === 'Нет') {
         tripSheet.getRange(lastRow, 1, 1, 20).setBackground('#ff4d4d');
       } else if (isWaitlist) {
         tripSheet.getRange(lastRow, 1, 1, 20).setBackground('#fb923c');
       }
    }
    
    // Подсветка жилья и бани (Зеленый цвет)
    const accVal = String(data.accommodation || "").toLowerCase();
    const extraVal = String(data.extraServices || "").toLowerCase();
    const hasAcc = accVal !== "" && accVal !== "-" && accVal !== "нет" && accVal !== "нет жилья";
    const hasExtra = extraVal !== "" && extraVal !== "-" && extraVal !== "нет";
    
    if (hasAcc || hasExtra) {
      const targetRange = tripSheet.getRange(lastRow, 12, 1, 2);
      targetRange.setBackground('#dcfce7').setFontWeight('bold').setBorder(true, true, true, true, null, null, "#10b981", SpreadsheetApp.BorderStyle.SOLID);
    }
    
    let otchetSheet = ss.getSheetByName('Отчеты');
    if (otchetSheet) {
      // Обновляем заголовки: Дата, Река, ФИО... Телефон
      const otchetHeaders = ["Дата", "Река", "ФИО", "Взрослые", "Дети", "Байдарки 2-мест", "Байдарки 3-мест", "Трансфер", "Собака", "Сумма", "Проживание", "Доп услуга", "Телефон"];
      otchetSheet.getRange(1, 1, 1, otchetHeaders.length).setValues([otchetHeaders]).setFontWeight('bold').setBackground('#dcfce7');
      
      const otchetRow = [
        sheetName,           // Дата (A)
        data.route,         // Река (B)
        data.name,          // ФИО (C)
        adultsCell,         // Взрослые (D)
        childrenCell,       // Дети (E)
        k2Cell,             // Байд 2 (F)
        k3Cell,             // Байд 3 (G)
        data.transfer,      // Трансфер (H)
        data.dog,           // Собака (I)
        totalSumCell,       // Сумма (J)
        data.accommodation || "-", // Проживание (K)
        data.extraServices || "-", // Доп услуга (L)
        "'" + data.phone    // Телефон (M)
      ];
      otchetSheet.appendRow(otchetRow);
    }

    let usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
      let uData = usersSheet.getDataRange().getValues();
      let uRow = -1;
      for (let i = 1; i < uData.length; i++) { if (uData[i][1].toString().replace(/\D/g,'').slice(-9) == cleanPhone.slice(-9)) { uRow = i + 1; break; } }
      if (uRow > 0) { 
        const newCount = (parseInt(uData[uRow-1][4]) || 0) + 1;
        usersSheet.getRange(uRow, 5).setValue(newCount);
        if (newCount >= 2) {
          usersSheet.getRange(uRow, 4).setValue('Постоянный');
          usersSheet.getRange(uRow, 1, 1, 5).setBackground('#dcfce7');
        }
      }
      else { usersSheet.appendRow([data.name, "'" + data.phone, "-", 'Новый', 1]); }
    }
    let adminMsg = "";
    if (isWaitlist) {
      adminMsg += `⚠️ <b>ЛИСТ ОЖИДАНИЯ!</b>\n\n`;
    } else if (prepayment > 0) {
      adminMsg += `💳 <b>ОПЛАТА ДОСТАВКИ ОБОРУДОВАНИЯ: ${prepayment} BYN!</b>\n\n`;
    } else {
      adminMsg += `🚀 <b>НОВАЯ ЗАЯВКА!</b>\n\n`;
    }
    
    adminMsg += `🆔 <b>ID Брони:</b> ${bookingId}\n`;
    adminMsg += `👤 <b>Клиент:</b> ${data.name}\n`;
    adminMsg += `📞 <b>Тел:</b> ${data.phone}\n`;
    adminMsg += `📅 <b>Сплав:</b> ${sheetName}\n`;
    adminMsg += `👥 <b>Участники:</b> ${data.adults} взр. + ${data.children} дет.\n`;
    adminMsg += `📱 <b>Telegram:</b> ${data.hasTelegram === 'Нет' ? '❌ НЕТ' : '✅ ЕСТЬ'}\n`;
    
    let kayakTextAdmin = data.kayakDisplay;
    if (isWaitlist) {
      kayakTextAdmin = "Закончились";
    } else if (data.kayakDisplay.includes("закончились")) {
      kayakTextAdmin = "2М закончились, предложена 3М";
    }
    adminMsg += `🚣 <b>Байдарки:</b> ${kayakTextAdmin}\n`;
    adminMsg += `🚌 <b>Проезд:</b> ${data.transfer === 'Да' ? 'ТРАНСФЕР' : 'СВОЁ АВТО'}\n\n`;
    
    if (finalComment) {
      adminMsg += `📝 <b>Комментарий:</b> ${finalComment}\n\n`;
    }
    
    if (data.accommodation && data.accommodation !== 'нет') {
      adminMsg += `🏠 <b>Жилье:</b> ${data.accommodation}\n`;
    }
    if (data.extraServices && data.extraServices !== 'нет') {
      adminMsg += `🧖 <b>Доп. услуги:</b> ${data.extraServices}\n`;
    }
    if (data.accommodation || data.extraServices) adminMsg += `\n`;

    if (prepayment > 0) {
      adminMsg += `💰 <b>К ОПЛАТЕ:</b> ${prepayment} BYN\n`;
      adminMsg += `💵 <b>Остаток на месте:</b> ${remaining} BYN\n\n`;
    }
    adminMsg += `💰 <b>ИТОГО:</b> ${totalSum} BYN`;

    sendTelegramMessage(CONFIG.ADMIN_CHAT_ID, adminMsg);
    
    // Обновляем статистику и логи
    updateStatistics(data.date, total, totalSum);
    logAction('Новая бронь', { name: data.name, phone: data.phone, total: total, sum: totalSum }, 'Успешно');
    
    // ВАЖНО: Обновляем занятость жилья в листе Сплавы
    try {
      const splav = ss.getSheetByName('Сплавы');
      const splavData = splav.getDataRange().getValues();
      const headers = splavData[0].map(h => String(h).trim().toLowerCase());
      const tz = ss.getSpreadsheetTimeZone();
      
      const dateIdx = headers.indexOf('дата');
      const timeIdx = headers.indexOf('время');
      
      for (let i = 1; i < splavData.length; i++) {
        const dStr = Utilities.formatDate(new Date(splavData[i][dateIdx]), ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");
        let tStr = "";
        const tVal = splavData[i][timeIdx];
        if (tVal instanceof Date) tStr = Utilities.formatDate(tVal, ss.getSpreadsheetTimeZone(), "HH:mm");
        else tStr = String(tVal).split(' ')[0];

        const splavRoute = String(splavData[i][headers.indexOf('маршрут')]).trim().toLowerCase();
        const targetRoute = String(data.route).trim().toLowerCase();

        if (dStr === data.date && tStr === data.time && splavRoute === targetRoute) {
          logAction('DEBUG_MATCH', `Match found at row ${i+1} for ${targetRoute}`, 'Success');
          // Нашли нужную строку, обновляем колонки Занято
          if (data.accommodation === 'Кемпер') {
            const colIdx = headers.indexOf('занято кемперов') + 1;
            if (colIdx > 0) splav.getRange(i + 1, colIdx).setValue((parseInt(splavData[i][colIdx-1]) || 0) + 1);
          } else if (data.accommodation === 'Типи') {
            const colIdx = headers.indexOf('занято типи') + 1;
            if (colIdx > 0) splav.getRange(i + 1, colIdx).setValue((parseInt(splavData[i][colIdx-1]) || 0) + 1);
          } else if (data.accommodation === 'Сафари-тент') {
            const colIdx = headers.indexOf('занято сафари') + 1;
            if (colIdx > 0) splav.getRange(i + 1, colIdx).setValue((parseInt(splavData[i][colIdx-1]) || 0) + 1);
          } else if (data.accommodation === 'Наша палатка') {
            let colIdx = headers.indexOf('занято палаток (ваших)') + 1;
            if (colIdx <= 0) colIdx = headers.indexOf('занято палаток (наших)') + 1;
            if (colIdx > 0) splav.getRange(i + 1, colIdx).setValue((parseInt(splavData[i][colIdx-1]) || 0) + 1);
          }
          
          if (data.extraServices && data.extraServices !== 'нет') {
            const colIdx = headers.indexOf('занято бань') + 1;
            if (colIdx > 0) splav.getRange(i + 1, colIdx).setValue((parseInt(splavData[i][colIdx-1]) || 0) + 1);
          }
          break;
        }
      }
    } catch (err) { logAction('Ошибка обновления Сплавов', err.toString(), 'Ошибка'); }

    // Обновляем остатки на листе Сплавы после новой брони (пересчет формул если есть)
    getAvailableDates();
    
    return { success: true, prepayment: prepayment, remaining: remaining, bookingId: bookingId };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function checkExistingBooking(phone, targetTripName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanPhone = phone.replace(/\D/g, "").slice(-9);
  const targetDate = targetTripName ? targetTripName.split('_')[0] : "";
  
  const allFutureBookings = [];
  const sheets = ss.getSheets();
  const ignore = ['Сплавы', 'Отчеты', 'Статистика', 'Users', 'Логи'];

  const findInSheet = (s) => {
    if (!s) return null;
    const data = s.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').replace(/\D/g, "").slice(-9) === cleanPhone) {
        const splavSheet = ss.getSheetByName('Сплавы');
        const splavData = splavSheet ? splavSheet.getDataRange().getValues() : [];
        const splavHeaders = splavData.length > 0 ? splavData[0].map(h => String(h).trim().toLowerCase()) : [];
        const gV = (r, n) => { const idx = splavHeaders.indexOf(n.toLowerCase()); return idx >= 0 ? r[idx] : ""; };

        let gps = "", program = "", pamyatka = "", route = "";
        for (let j = 1; j < splavData.length; j++) {
           const dStr = Utilities.formatDate(new Date(splavData[j][0]), ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");
           if (s.getName().indexOf(dStr) !== -1) {
             gps = String(gV(splavData[j], 'Координаты') || gV(splavData[j], 'GPS') || '');
             program = String(gV(splavData[j], 'Программа') || '');
             pamyatka = String(gV(splavData[j], 'Памятка') || '');
             route = String(gV(splavData[j], 'Маршрут') || '');
             break;
           }
        }
        return { found: true, date: s.getName().split(' ')[0], time: s.getName().split(' ')[1] || "", tripName: s.getName(), name: data[i][0], bookingId: data[i][16], prepayment: data[i][17], remaining: data[i][18], status: data[i][19], adults: data[i][4], children: data[i][5], gps: gps, program: program, pamyatka: pamyatka, route: route };
      }
    }
    return null;
  };

  for (let s of sheets) {
    if (ignore.indexOf(s.getName()) !== -1) continue;
    const res = findInSheet(s);
    if (res) allFutureBookings.push(res);
  }

  if (allFutureBookings.length === 0) return { found: false };

  const sameDayMatch = allFutureBookings.find(b => b.date === targetDate);
  if (sameDayMatch) {
    return { ...sameDayMatch, type: 'same_day', allCount: allFutureBookings.length };
  }

  return { ...allFutureBookings[0], type: 'other_day', allCount: allFutureBookings.length };
}

/**
 * Перекрашивает все строчки клиента в синий во всех активных листах
 */
function repaintMultiBookings(phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanPhone = phone.replace(/\D/g, "").slice(-9);
  const sheets = ss.getSheets();
  const ignore = ['Сплавы', 'Отчеты', 'Статистика', 'Users', 'Логи'];
  const blueColor = '#a4c2f4'; // Яркий насыщенный голубой

  sheets.forEach(s => {
    if (ignore.indexOf(s.getName()) !== -1) return;
    const data = s.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').replace(/\D/g, "").slice(-9) === cleanPhone) {
        s.getRange(i + 1, 1, 1, 20).setBackground(blueColor);
      }
    }
  });
}

function cancelBooking(phone, tripName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanPhone = phone.replace(/\D/g, "").slice(-9);
  logAction('Запрос на отмену', {phone: cleanPhone, trip: tripName}, 'Начало поиска');
  
  const sheet = ss.getSheetByName(tripName);
  if (!sheet) {
    logAction('Ошибка отмены', `Лист ${tripName} не найден`, 'Ошибка');
    // Не выходим, пробуем чистить другие листы
  }
  
  const data = (sheet) ? sheet.getDataRange().getValues() : [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').replace(/\D/g, "").slice(-9) === cleanPhone) {
      const row = data[i];
      const adults = parseInt(row[4]) || 0;
      const children = parseInt(row[5]) || 0;
      const k2 = parseInt(row[6]) || 0;
      const k3 = parseInt(row[7]) || 0;
      const bookingId = row[16] || "—";
      
      let kayakText = [];
      if (k2 > 0) kayakText.push(k2 + " шт 2-местных");
      if (k3 > 0) kayakText.push(k3 + " шт 3-местных");
      
      let cancelMsg = `❌ <b>ОТМЕНА ЗАЯВКИ!</b>\n\n`;
      cancelMsg += `🆔 <b>ID Брони:</b> ${bookingId}\n`;
      cancelMsg += `👤 <b>Клиент:</b> ${row[0]}\n`;
      cancelMsg += `📞 <b>Тел:</b> ${row[1]}\n`;
      cancelMsg += `👥 <b>Участники:</b> ${adults + children} чел. (${adults} взр. + ${children} дет.)\n`;
      cancelMsg += `📅 <b>Сплав:</b> ${tripName}\n`;
      cancelMsg += `🚣 <b>Байдарки:</b> ${kayakText.join(' + ') || "не указаны"}\n`;
      cancelMsg += `🚌 <b>Проезд:</b> ${row[8] > 0 ? 'ТРАНСФЕР' : 'СВОЁ АВТО'}\n`;
      
      sendTelegramMessage(CONFIG.ADMIN_CHAT_ID, cancelMsg);

      // ШАГ 1: Удаляем из листа даты
      try {
        sheet.deleteRow(i + 1);
        logAction('Отмена (Шаг 1)', { trip: tripName, name: row[0] }, 'Успешно удалено из листа даты');
      } catch (e) {
        logAction('Ошибка (Шаг 1)', e.toString(), 'Не удалось удалить из листа даты');
      }
      
      // ШАГ 2: Удаляем из отчетов
      try {
        const otchet = ss.getSheetByName('Отчеты');
        if (otchet) {
          const oData = otchet.getDataRange().getValues();
          const clientName = String(row[0]).trim().toLowerCase();
          const bId = String(row[16] || '').trim();
          
          let otchetDeleted = false;
          for (let j = oData.length - 1; j >= 1; j--) {
            const oTripRaw = oData[j][0];
            const oTripStr = oTripRaw instanceof Date ? Utilities.formatDate(oTripRaw, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm") : String(oTripRaw).trim();
            const oPhone = String(oData[j][1] || '').replace(/\D/g, "").slice(-9);
            const oName = String(oData[j][3] || oData[j][2] || '').trim().toLowerCase();
            
            // Сравнение: по Дате (начало строки) + (Телефон ИЛИ Имя)
            const dateMatch = oTripStr.indexOf(tripName.trim()) !== -1;
            const phoneMatch = oPhone === cleanPhone;
            const nameMatch = oName === clientName;
            
            if ((dateMatch && (phoneMatch || nameMatch)) || (phoneMatch && nameMatch)) {
              otchet.deleteRow(j + 1); 
              otchetDeleted = true;
              logAction('Отмена (Шаг 2)', { trip: tripName, name: row[0] }, 'Успешно удалено из Отчетов');
              break;
            }
          }
          // ФОЛБЭК: Если по дате не нашли, ищем просто по телефону (самую последнюю запись)
          if (!otchetDeleted) {
            for (let j = oData.length - 1; j >= 1; j--) {
              if (String(oData[j][1] || '').replace(/\D/g, "").slice(-9) === cleanPhone) {
                otchet.deleteRow(j + 1);
                otchetDeleted = true;
                logAction('Отмена (Шаг 2 - Фолбэк)', { phone: cleanPhone }, 'Удалено по номеру телефона');
                break;
              }
            }
          }
          if (!otchetDeleted) logAction('Предупреждение (Шаг 2)', `Не удалось найти запись ни одним способом`, 'Пропущено');
        }
      } catch (e) {
        logAction('Ошибка (Шаг 2)', e.toString(), 'Сбой при удалении из Отчетов');
      }

      // ШАГ 3: Полностью удаляем из Users
      try {
        const usersSheet = ss.getSheetByName('Users');
        if (usersSheet) {
          const uData = usersSheet.getDataRange().getValues();
          for (let k = uData.length - 1; k >= 1; k--) {
            if (String(uData[k][1] || '').replace(/\D/g, "").slice(-9) === cleanPhone) {
              usersSheet.deleteRow(k + 1);
              logAction('Отмена (Шаг 3)', { phone: cleanPhone }, 'Строка полностью удалена из Users');
              break;
            }
          }
        }
      } catch (e) {
        logAction('Ошибка (Шаг 3)', e.toString(), 'Сбой при удалении из Users');
      }

      // ШАГ 4: Удаляем из Статистика
      try {
        const statsSheet = ss.getSheetByName('Статистика') || ss.getSheetByName('Статистика ');
        if (statsSheet) {
          const sData = statsSheet.getDataRange().getValues();
          let statsDeleted = false;
          for (let m = sData.length - 1; m >= 1; m--) {
            const sTrip = String(sData[m][0]).trim();
            const sPhone = String(sData[m][1] || '').replace(/\D/g, "").slice(-9);
            const sName = String(sData[m][3] || sData[m][2] || '').trim().toLowerCase();
            
            if (sTrip.indexOf(tripName.trim()) !== -1 && (sPhone === cleanPhone || sName === clientName)) {
              statsSheet.deleteRow(m + 1);
              statsDeleted = true;
              logAction('Отмена (Шаг 4)', { trip: tripName }, 'Строка удалена из Статистики');
              break;
            }
          }
          if (!statsDeleted) logAction('Предупреждение (Шаг 4)', `Запись не найдена в Статистике`, 'Пропущено');
        } else {
          logAction('Предупреждение (Шаг 4)', `Лист "Статистика" не найден`, 'Пропущено');
        }
      } catch (e) {
        logAction('Ошибка (Шаг 4)', e.toString(), 'Сбой в Статистике');
      }

      // ШАГ 5: Освобождаем жилье и баню в листе Сплавы
      try {
        const splav = ss.getSheetByName('Сплавы');
        if (splav) {
          const splavData = splav.getDataRange().getValues();
          const splavHeaders = splavData[0].map(h => String(h).trim().toLowerCase());
          const dateIdx = splavHeaders.indexOf('дата');
          const timeIdx = splavHeaders.indexOf('время');
          
          const tripDateStr = tripName.split(' ')[0];
          const tripTimeStr = tripName.split(' ')[1] || "";
          
          for (let l = 1; l < splavData.length; l++) {
            const dStr = Utilities.formatDate(new Date(splavData[l][dateIdx]), ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");
            let tStr = "";
            const tVal = splavData[l][timeIdx];
            if (tVal instanceof Date) tStr = Utilities.formatDate(tVal, ss.getSpreadsheetTimeZone(), "HH:mm");
            else tStr = String(tVal).split(' ')[0];

            if (dStr === tripDateStr && tStr === tripTimeStr) {
              const accVal = String(row[11] || "").trim();
              const accValLC = accVal.toLowerCase();
              const extraVal = String(row[12] || "").trim();
              
              if (accVal === 'Кемпер') {
                const colIdx = splavHeaders.indexOf('занято кемперов') + 1;
                if (colIdx > 0) splav.getRange(l + 1, colIdx).setValue(Math.max(0, (parseInt(splavData[l][colIdx-1]) || 0) - 1));
              } else if (accVal === 'Типи') {
                const colIdx = splavHeaders.indexOf('занято типи') + 1;
                if (colIdx > 0) splav.getRange(l + 1, colIdx).setValue(Math.max(0, (parseInt(splavData[l][colIdx-1]) || 0) - 1));
              } else if (accVal === 'Сафари-тент') {
                const colIdx = splavHeaders.indexOf('занято сафари') + 1;
                if (colIdx > 0) splav.getRange(l + 1, colIdx).setValue(Math.max(0, (parseInt(splavData[l][colIdx-1]) || 0) - 1));
              } else if (accVal === 'Наша палатка') {
                let colIdx = splavHeaders.indexOf('занято палаток (ваших)') + 1;
                if (colIdx <= 0) colIdx = splavHeaders.indexOf('занято палаток (наших)') + 1;
                if (colIdx > 0) splav.getRange(l + 1, colIdx).setValue(Math.max(0, (parseInt(splavData[l][colIdx-1]) || 0) - 1));
              }
              
              if (extraVal && extraVal.toLowerCase().includes('баня')) {
                const colIdx = splavHeaders.indexOf('занято бань') + 1;
                if (colIdx > 0) splav.getRange(l + 1, colIdx).setValue(Math.max(0, (parseInt(splavData[l][colIdx-1]) || 0) - 1));
              }
              logAction('Отмена (Шаг 5)', { trip: tripName }, 'Ресурсы освобождены в Сплавах');
              break;
            }
          }
        }
      } catch (e) {
        logAction('Ошибка (Шаг 5)', e.toString(), 'Сбой при освобождении ресурсов');
      }

      getAvailableDates();
      return { success: true };
    }
  }
  logAction('Ошибка отмены', { phone: phone, trip: tripName }, 'Бронь не найдена');
  return { success: false };
}

function sendTelegramMessage(chatId, text, replyMarkup) {
  try {
    const url = `https://api.telegram.org/bot${CONFIG.BOT_TOKEN}/sendMessage`;
    const payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
    if (replyMarkup) payload.reply_markup = replyMarkup;
    UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  } catch (e) {}
}

function doPost(e) {
  try {
    const contents = e.postData.contents;
    logAction('Входящий запрос', contents, 'Получено');
    const update = JSON.parse(contents);
    if (update.message) {
      if (update.message.contact) handleContact(update.message);
      else if (update.message.photo) handlePhoto(update.message);
      else handleTelegramMessage(update.message);
    } else if (update.callback_query) {
      handleCallback(update.callback_query);
    }
  } catch (err) {}
}

function handleTelegramMessage(msg) {
  const chatId = String(msg.chat.id);
  const text = (msg.text || "").trim();
  const isAdmin = (chatId === CONFIG.ADMIN_CHAT_ID);

  if (text.toLowerCase().startsWith("/start")) {
    if (isAdmin) {
      sendTelegramMessage(chatId, "👋 <b>Панель Администратора</b>\n\nПришлите номер телефона клиента (9 цифр, например 291234567), чтобы найти брони и управлять ими.");
    } else {
      const parts = text.split(' ');
      if (parts.length > 1 && parts[1].startsWith("PAY_")) {
        const bookingId = parts[1].replace("PAY_", "");
        sendTelegramMessage(chatId, "🔎 <b>Ищу вашу бронь...</b>");
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = ss.getSheets();
        let found = false;
        
        for (let s of sheets) {
          if (s.isSheetHidden()) continue;
          const sName = s.getName();
          if (['Сплавы', 'Статистика', 'Отчеты', 'Users', 'Логи'].indexOf(sName) !== -1 || sName.toLowerCase().includes('архив')) continue;
          
          const data = s.getDataRange().getValues();
          for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][16] || '') === bookingId) {
              const bId = data[i][16] || "";
              const prepayment = data[i][17] || 0;
              const remaining = data[i][18] || 0;
              const status = data[i][19] || "";
              
              const phone = String(data[i][1] || '').replace(/\D/g, '').slice(-9);
              let usersSheet = ss.getSheetByName('Users');
              if (usersSheet) {
                let uData = usersSheet.getDataRange().getValues();
                if (usersSheet.getLastColumn() < 6) usersSheet.getRange(1, 6).setValue("Chat ID");
                for (let k = 1; k < uData.length; k++) {
                  if (String(uData[k][1] || '').replace(/\D/g, '').slice(-9) === phone) {
                    usersSheet.getRange(k + 1, 6).setValue(chatId);
                    usersSheet.getRange(k + 1, 7).setValue(bookingId); // Запоминаем, что именно этот ID сейчас оплачивается
                    break;
                  }
                }
              }

              if (status.toString().toLowerCase().includes("оплачен")) {
                sendTelegramMessage(chatId, `🚣 <b>Бронирование Mazaysplav</b>\n\n👋 Привет, <b>${data[i][0]}</b>!\nВы уже оплачивали доставку оборудования и она принята. ✅\n(Найдено в листе: ${sName})`);
              } else {
                const payMsg = `🚣 <b>Бронирование Mazaysplav</b>\n\n👋 Привет, <b>${data[i][0]}!</b>\nДля подтверждения бронирования необходимо оплатить доставку оборудования:\n\n🆔 <b>Ваш ID:</b> ${bId}\n💰 <b>Оплата доставки:</b> ${prepayment} BYN\n💵 <b>Остаток на месте:</b> ${remaining} BYN\n\n1️⃣ <b>Ссылка для оплаты:</b> https://ecom.alfabank.by/sc/JUUtIHIhTRMGkgJp\n2️⃣ <b>После оплаты</b> пришлите <b>фото чека</b> прямо в этот чат. ✅`;
                sendTelegramMessage(chatId, payMsg);
              }
              found = true;
              break;
            }
          }
          if (found) break;
        }
        if (!found) sendTelegramMessage(chatId, "❌ Бронь с таким ID не найдена. Попробуйте найти по номеру телефона.");
      } else {
        const keyboard = { keyboard: [[{ text: "📱 Найти мою бронь", request_contact: true }]], resize_keyboard: true, one_time_keyboard: true };
        sendTelegramMessage(chatId, "👋 Привет! Нажмите кнопку ниже 👇", keyboard);
      }
    }
    return;
  }

  // Админ поиск по номеру
  const cleanText = text.replace(/\D/g, '');
  if (isAdmin && cleanText.length >= 9) {
    const phone = cleanText.slice(-9);
    sendTelegramMessage(chatId, "🔎 <b>Ищу бронирование по номеру ..."+phone+"</b>");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let found = false;
    
    for (let s of sheets) {
      if (s.isSheetHidden()) continue;
      const sName = s.getName();
      if (['Сплавы', 'Статистика', 'Отчеты', 'Users', 'Логи'].indexOf(sName) !== -1 || sName.toLowerCase().includes('архив')) continue;
      
      const data = s.getDataRange().getValues();
      // Ищем с конца (самые свежие записи)
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1] || '').replace(/\D/g, '').slice(-9) === phone) {
          const info = `🔎 <b>Бронь найдена:</b>\n👤 Клиент: ${data[i][0]}\n📅 Сплав: ${sName}\n👥 Участники: ${data[i][4]} взр. + ${data[i][5]} дет.\n🏠 Жилье: ${data[i][11] || "-"}\n🧖 Услуги: ${data[i][12] || "-"}`;
          const keyboard = {
            inline_keyboard: [
              [{ text: "❌ Отменить бронь", callback_data: `CANCEL|${phone}|${sName}` }],
              [{ text: "📅 Перенести на другую дату", callback_data: `RESCH_LIST|${phone}|${sName}` }]
            ]
          };
          sendTelegramMessage(chatId, info, keyboard);
          found = true;
          break;
        }
      }
      if (found) break;
    }
    if (!found) sendTelegramMessage(chatId, "❌ Бронь с таким номером не найдена.");
    return;
  }
}

function handleCallback(query) {
  const chatId = query.message.chat.id;
  const data = query.data;
  const [action, p1, p2, p3] = data.split("|");

  if (action === "CANCEL") {
    const res = cancelBooking(p1, p2);
    if (res.success) sendTelegramMessage(chatId, `✅ Бронь на <b>${p2}</b> удалена.`);
  }

  if (action === "RESCH_LIST") {
    const phone = p1;
    const oldTrip = p2;
    const trips = getAvailableDates();
    const buttons = trips.filter(t => (t.date + " " + t.time) !== oldTrip).map(t => [
      { text: `📅 ${t.date} ${t.time}`, callback_data: `RESCH_DO|${phone}|${oldTrip}|${t.date} ${t.time}` }
    ]);
    if (buttons.length === 0) {
      sendTelegramMessage(chatId, "❌ Нет доступных дат для переноса.");
    } else {
      sendTelegramMessage(chatId, "📅 Выберите новую дату для переноса:", { inline_keyboard: buttons.slice(0, 10) });
    }
  }

  if (action === "RESCH_DO") {
    const phone = p1;
    const oldTrip = p2;
    const newTrip = p3;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const oldSheet = ss.getSheetByName(oldTrip);
    if (!oldSheet) return;
    const oldData = oldSheet.getDataRange().getValues();
    let bookingData = null;
    for (let i = 1; i < oldData.length; i++) {
      if (String(oldData[i][1] || '').replace(/\D/g, '').slice(-9) === phone.slice(-9)) {
        bookingData = oldData[i];
        break;
      }
    }
    if (bookingData) {
      cancelBooking(phone, oldTrip);
      const [newDate, newTime] = newTrip.split(' ');
      const trips = getAvailableDates();
      const newTripObj = trips.find(t => t.date === newDate && t.time === newTime);
      
      const adults = parseInt(bookingData[4]) || 0;
      const children = parseInt(bookingData[5]) || 0;
      const total = adults + children;
      const needTransfer = bookingData[8] > 0;
      
      // Проверяем, влезет ли бронь на новую дату
      let isNewWaitlist = false;
      let k2Req = 0, k3Req = 0;
      if (total === 1) isNewWaitlist = true;
      else {
        if (total % 2 === 0) { k2Req = total / 2; k3Req = 0; }
        else { k3Req = 1; k2Req = (total - 3) / 2; }
        if (k2Req > newTripObj.rem2M || k3Req > newTripObj.rem3M) isNewWaitlist = true;
      }
      if (needTransfer && total > newTripObj.remTransfer) isNewWaitlist = true;

      const accType = "-"; // Сбрасываем жилье при переносе
      const extraType = "-"; // Сбрасываем услуги при переносе
      
      const priceA = needTransfer ? (newTripObj.priceAdultWith || 0) : (newTripObj.priceAdultNo || 0);
      const priceC = needTransfer ? (newTripObj.priceChildWith || 0) : (newTripObj.priceChildNo || 0);
      const newTotalSum = (adults * priceA) + (children * priceC);

      const formData = {
        name: bookingData[0], phone: phone,
        date: newDate, time: newTime, adults: adults, children: children,
        dog: bookingData[9], comment: bookingData[10], transfer: needTransfer ? "Да" : "Нет",
        k2: k2Req, k3: k3Req, route: newTripObj.route, type: newTripObj.type,
        dayType: newTripObj.dayType, unitPriceAdult: priceA,
        unitPriceChild: priceC,
        kayakDisplay: isNewWaitlist ? "Места закончились (Лист ожидания)" : ((k2Req > 0 ? k2Req + " шт 2-местных " : "") + (k3Req > 0 ? k3Req + " шт 3-местных" : "")),
        isWaitlist: isNewWaitlist,
        accommodation: accType,
        extraServices: extraType,
        sum: newTotalSum
      };
      processBooking(formData);
      sendTelegramMessage(chatId, `✅ Бронь успешно перенесена на <b>${newTrip}</b>${isNewWaitlist ? " (в ЛИСТ ОЖИДАНИЯ ⚠️)" : ""}.\n\n💰 <b>Новая сумма:</b> ${newTotalSum} BYN\n⚠️ <b>ВНИМАНИЕ:</b> Жилье и доп. услуги НЕ перенесены. Проверьте наличие мест на новую дату и добавьте их вручную, если нужно.`);
    }
  }

  UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.BOT_TOKEN}/answerCallbackQuery`, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify({ callback_query_id: query.id }), muteHttpExceptions: true
  });
}

function handleContact(msg) {
  const chatId = msg.chat.id;
  const cleanPhone = msg.contact.phone_number.replace(/\D/g, '').slice(-9);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (let s of sheets) {
    if (s.isSheetHidden()) continue;
    const sName = s.getName();
    if (['Сплавы', 'Статистика', 'Отчеты', 'Users', 'Логи'].indexOf(sName) !== -1 || sName.toLowerCase().includes('архив')) continue;
    const data = s.getDataRange().getValues();
    // Ищем с конца (самые свежие записи)
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1] || '').replace(/\D/g, '').slice(-9) === cleanPhone) {
        let usersSheet = ss.getSheetByName('Users');
        if (usersSheet) {
          let uData = usersSheet.getDataRange().getValues();
          if (usersSheet.getLastColumn() < 6) usersSheet.getRange(1, 6).setValue("Chat ID");
          for (let k = 1; k < uData.length; k++) { if (String(uData[k][1] || '').replace(/\D/g, '').slice(-9) === cleanPhone) { usersSheet.getRange(k + 1, 6).setValue(chatId); break; } }
        }
        const bId = data[i][16] || "";
        const prepayment = data[i][17] || 0;
        const remaining = data[i][18] || 0;
        const status = data[i][19] || "";
        
        if (status.toString().toLowerCase().includes("оплачен")) {
          sendTelegramMessage(chatId, `🚣 <b>Бронирование Mazaysplav</b>\n\n👋 Привет, <b>${data[i][0]}</b>!\nВы уже оплачивали доставку оборудования и она принята. ✅\n(Найдено в листе: ${sName})`);
          return;
        }
        
        const payMsg = `🚣 <b>Бронирование Mazaysplav</b>\n\n👋 Привет, <b>${data[i][0]}</b>!\nДля подтверждения бронирования необходимо оплатить доставку оборудования:\n\n🆔 <b>Ваш ID:</b> ${bId}\n💰 <b>Оплата доставки:</b> ${prepayment} BYN\n💵 <b>Остаток на месте:</b> ${remaining} BYN\n\n1️⃣ <b>Ссылка для оплаты:</b> https://ecom.alfabank.by/sc/JUUtIHIhTRMGkgJp\n2️⃣ <b>После оплаты</b> пришлите <b>фото чека</b> прямо в этот чат. ✅`;
        sendTelegramMessage(chatId, payMsg);
        return;
      }
    }
  }
  sendTelegramMessage(chatId, `❌ Бронь не найдена.`);
}

function handlePhoto(msg) {
  const chatId = msg.chat.id;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  let phone = "";
  let activeId = "";
  if (usersSheet) {
    const uData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      if (uData[i][5] && uData[i][5].toString() === chatId.toString()) {
        phone = String(uData[i][1] || '').replace(/\D/g, '').slice(-9);
        activeId = String(uData[i][6] || ''); // Получаем ID, который клиент сейчас оплачивает
        break;
      }
    }
  }
  if (!phone) { sendTelegramMessage(chatId, "⚠️ Сначала нажмите «Найти бронь»."); return; }
  
  const sheets = ss.getSheets();
  let found = false;

  // Сначала ищем по точному активному ID (если он есть)
  if (activeId) {
    for (let s of sheets) {
      if (['Сплавы', 'Статистика', 'Отчеты', 'Users', 'Логи'].indexOf(s.getName()) !== -1) continue;
      const data = s.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][16] || '') === activeId) {
          processPhotoMatch(s, i, data[i], chatId);
          found = true; break;
        }
      }
      if (found) break;
    }
  }

  // Если по ID не нашли или его не было — ищем по телефону (с конца листа, самые свежие)
  if (!found) {
    for (let s of sheets) {
      if (['Сплавы', 'Статистика', 'Отчеты', 'Users', 'Логи'].indexOf(s.getName()) !== -1) continue;
      const data = s.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1] || '').replace(/\D/g, '').slice(-9) === phone) {
          processPhotoMatch(s, i, data[i], chatId);
          found = true; break;
        }
      }
      if (found) break;
    }
  }
}

function processPhotoMatch(sheet, rowIndex, rowData, chatId) {
  sheet.getRange(rowIndex + 1, 18).setBackground('#dcfce7');
  sheet.getRange(rowIndex + 1, 20).setValue('Оплачено');
  const remaining = rowData[18] || 0;
  sendTelegramMessage(chatId, `✅ <b>Чек получен, спасибо!</b> Мы подтверждаем вашу бронь. На месте вам осталось оплатить <b>${remaining}</b> BYN. Мы с вами свяжемся за несколько дней до сплава.`); 
  
  const adminCheckMsg = `✅ <b>ПОЛУЧЕН ЧЕК!</b>\n👤 ${rowData[0]}\n📅 ${sheet.getName()}\n👥 Участники: ${rowData[4]} взр. + ${rowData[5]} дет.\n💰 Оплата доставки: ${rowData[17]} BYN`;
  sendTelegramMessage(CONFIG.ADMIN_CHAT_ID, adminCheckMsg);
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Отчеты');
    if (!sheet) {
      // Пытаемся найти лист без учета регистра
      const allSheets = ss.getSheets();
      sheet = allSheets.find(s => s.getName().trim().toLowerCase() === 'отчеты');
    }
    
    if (!sheet) return { error: 'Лист "Отчеты" не найден' };
    
    const data = sheet.getDataRange().getValues();
    const stats = {
      revenue: { total: 0, season: 0, month: 0, week: 0 },
      participants: { total: 0, adults: 0, kids: 0, bookings: 0 },
      logistics: { transfer: 0, auto: 0 },
      rivers: {},
      accommodation: {},
      monthly: {},
      tripsCount: 0
    };

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Начало сегодняшнего дня
    
    const splavSheet = ss.getSheetByName('Сплавы');
    if (splavSheet) {
      const sData = splavSheet.getDataRange().getValues();
      const sHeaders = sData[0].map(h => String(h).trim().toLowerCase());
      const dateIdx = sHeaders.indexOf('дата');
      if (dateIdx !== -1) {
        stats.tripsCount = sData.slice(1).filter(row => {
          const d = row[dateIdx];
          return d instanceof Date && d < today; // Только завершенные (вчера и ранее)
        }).length;
      }
    }

    if (data.length < 2) return stats; // Если только шапка или пусто - возвращаем нули
    
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const rows = data.slice(1);
    stats.participants.bookings = rows.length;
    
    // Динамический поиск индексов колонок
    const idx = {
      date: headers.indexOf('дата'),
      river: headers.indexOf('река'),
      adults: headers.indexOf('взрослые'),
      kids: headers.indexOf('дети'),
      transfer: headers.indexOf('трансфер'),
      sum: headers.indexOf('сумма'),
      accommodation: headers.indexOf('проживание'),
      extra: headers.indexOf('доп услуга')
    };

    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();
    
    const tempDate = new Date(now);
    const day = tempDate.getDay();
    const diff = tempDate.getDate() - day + (day === 0 ? -6 : 1);
    const startOfWeek = new Date(tempDate.setDate(diff));
    startOfWeek.setHours(0,0,0,0);
    
    const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
    for (let m = 0; m <= 11; m++) {
      const key = (m + 1).toString().padStart(2, '0') + "." + currentYear;
      stats.monthly[key] = { name: monthNames[m], trips: 0, people: 0, revenue: 0 };
    }
    
    if (rows.length === 0) return stats;

    // Сбрасываем bookings, будем считать только подходящие по дате
    stats.participants.bookings = 0;

    rows.forEach(row => {
      let rawDate = idx.date !== -1 ? row[idx.date] : null;
      if (!rawDate) return;
      
      let date;
      if (rawDate instanceof Date) {
        date = rawDate;
      } else {
        const dStr = String(rawDate).split(' ')[0];
        const p = dStr.split('.');
        if (p.length === 3) date = new Date(p[2], p[1]-1, p[0]);
      }
      if (!date || isNaN(date.getTime())) return;

      // ФИЛЬТР: Пропускаем сегодня и будущие сплавы
      if (date >= today) return;

      stats.participants.bookings++; // Считаем только завершенные брони

      let river = idx.river !== -1 ? String(row[idx.river] || "").trim() : "—";
      if (river.indexOf('+') !== -1 && idx.river !== -1) {
        river = String(row[idx.river + 1] || "").trim();
      }
      
      const adults = idx.adults !== -1 ? Number(row[idx.adults]) || 0 : 0;
      const kids = idx.kids !== -1 ? Number(row[idx.kids]) || 0 : 0;
      const isTransfer = idx.transfer !== -1 ? String(row[idx.transfer]).toLowerCase() === 'да' : false;
      const sum = idx.sum !== -1 ? Number(row[idx.sum]) || 0 : 0;
      
      let accType = idx.accommodation !== -1 ? String(row[idx.accommodation] || "-").trim() : "-";
      if (!isNaN(accType) && accType !== "" && accType !== "-" && idx.accommodation !== -1) {
         accType = String(row[idx.accommodation + 1] || "-").trim();
      }

      let extraSrv = idx.extra !== -1 ? String(row[idx.extra] || "-").trim() : "-";
      
      // Финансы
      stats.revenue.total += sum;
      if (date.getFullYear() === currentYear) stats.revenue.season += sum;
      if (date.getFullYear() === currentYear && date.getMonth() === currentMonth) stats.revenue.month += sum;
      if (date >= startOfWeek) stats.revenue.week += sum;
      
      // Люди
      stats.participants.total += (adults + kids);
      stats.participants.adults += adults;
      stats.participants.kids += kids;
      if (isTransfer) stats.logistics.transfer++; else stats.logistics.auto++;
      
      // Реки
      if (river && river !== "Река" && river !== "—") stats.rivers[river] = (stats.rivers[river] || 0) + 1;
      
      // Жилье и Услуги
      const recordType = (type) => {
        if (type && type !== "-" && type !== "нет" && type !== "нет услуг") {
          stats.accommodation[type] = (stats.accommodation[type] || 0) + 1;
        }
      };
      recordType(accType);
      recordType(extraSrv);
      
      const monthKey = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MM.yyyy");
      if (stats.monthly[monthKey]) {
        stats.monthly[monthKey].trips++;
        stats.monthly[monthKey].people += (adults + kids);
        stats.monthly[monthKey].revenue += sum;
      }
    });
    
    return stats;
  } catch (err) {
    return { error: err.toString() };
  }
}

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (СТАТИСТИКА И ЛОГИ)
// ============================================
function logAction(action, data, result) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logs = ss.getSheetByName('Логи');
    if (!logs) {
      logs = ss.insertSheet('Логи');
      logs.appendRow(['Timestamp', 'Действие', 'Данные', 'Результат']);
      logs.getRange('A1:D1').setFontWeight('bold').setBackground('#9e9e9e').setFontColor('#ffffff');
    }
    
    // Восстановление заголовков если они пропали или лист был очищен
    if (logs.getLastRow() === 0 || logs.getRange(1, 1).getValue() !== 'Timestamp') {
      if (logs.getLastRow() > 0) logs.insertRowBefore(1);
      logs.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Действие', 'Данные', 'Результат']]).setFontWeight('bold').setBackground('#9e9e9e').setFontColor('#ffffff');
    }
    
    let readableData = "";
    if (typeof data === 'object') {
      if (data.name) {
        readableData = `Клиент: ${data.name}, Тел: ${data.phone}, Итого: ${data.sum || data.totalSum || '0'} BYN`;
      } else {
        readableData = JSON.stringify(data);
      }
    } else {
      readableData = String(data);
      if (action === 'Входящий запрос' && readableData.includes('{')) {
        try {
          const upd = JSON.parse(readableData);
          const msg = upd.message || upd.callback_query;
          if (msg) {
            const user = msg.from ? (msg.from.username || msg.from.first_name) : 'user';
            const content = msg.text || msg.data || (msg.photo ? 'Фото' : (msg.contact ? 'Контакт' : 'Медиа'));
            readableData = `От ${user}: ${content}`;
          }
        } catch (e) {}
      }
    }
    
    // Вставляем новую строку ВСЕГДА ВТОРОЙ (сразу после заголовка)
    // Это решает проблему пустых строк и делает лог удобнее (новые записи сверху)
    logs.insertRowAfter(1);
    logs.getRange(2, 1, 1, 4).setValues([[new Date(), action, readableData, result]]);
    
  } catch (e) {}
}

function updateStatistics(date, peopleCount, sum) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let stats = ss.getSheetByName('Статистика');
    if (!stats) {
      stats = ss.insertSheet('Статистика');
      stats.appendRow(['Дата', 'Всего броней', 'Всего людей', 'Выручка', 'Средний чек']);
      stats.getRange('A1:E1').setFontWeight('bold').setBackground('#fbbc04').setFontColor('#ffffff');
    }
    const data = stats.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      const rowDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowDate === date) {
        const bookings = (data[i][1] || 0) + 1;
        const people = (data[i][2] || 0) + peopleCount;
        const revenue = (data[i][3] || 0) + sum;
        const avgCheck = revenue / bookings;

        stats.getRange(i + 1, 2).setValue(bookings);
        stats.getRange(i + 1, 3).setValue(people);
        stats.getRange(i + 1, 4).setValue(revenue);
        stats.getRange(i + 1, 5).setValue(avgCheck);
        found = true;
        break;
      }
    }
    if (!found) {
      stats.appendRow([date, 1, peopleCount, sum, sum]);
    }
  } catch (e) {
    logAction('Ошибка статистики', e.toString(), 'Ошибка');
  }
}
function setupWebhook() {
  // ЕСЛИ АВТОМАТИЧЕСКИ НЕ РАБОТАЕТ, ВСТАВЬТЕ ССЫЛКУ СЮДА ВРУЧНУЮ:
  let manualUrl = "https://script.google.com/macros/s/AKfycbxGSupSyj1AxsE2713alKD4a3TgiwZR4Kw-t4rdHjOeM29InFO2S8JZH7fmXPyM5Mc/exec"; 
  let url = manualUrl || ScriptApp.getService().getUrl();
  
  if (!url || url.indexOf('dev') !== -1) {
    Logger.log("❌ ОШИБКА: Не удалось получить URL. Пожалуйста, сделайте следующее:");
    Logger.log("1. Нажмите «Развернуть» -> «Управление развертываниями»");
    Logger.log("2. Скопируйте URL веб-приложения (/exec)");
    Logger.log("3. Вставьте его в код функции setupWebhook в переменную manualUrl");
    return;
  }

  const tgUrl = `https://api.telegram.org/bot${CONFIG.BOT_TOKEN}/setWebhook?url=${url}`;
  const response = UrlFetchApp.fetch(tgUrl);
  Logger.log("✅ БОТ НАСТРОЕН! Ссылка: " + url);
  Logger.log("Telegram ответил: " + response.getContentText());
}

/**
 * ФУНКЦИИ ДЛЯ WEB-ДАШБОРДА
 */

function onOpen() {
  // Удаляем старый дашборд при открытии, если он еще есть
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldDash = ss.getSheetByName('Дашборд');
  if (oldDash) ss.deleteSheet(oldDash);
  const tech = ss.getSheetByName('TechData');
  if (tech) ss.deleteSheet(tech);

  SpreadsheetApp.getUi()
    .createMenu('🚀 MAZAY ANALYTICS')
    .addItem('Открыть Web-Дашборд', 'showWebDashboard')
    .addToUi();
}

function showWebDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('WebDashboard')
    .setWidth(1150)
    .setHeight(750)
    .setTitle('Mazaysplav Analytics Pro');
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Аналитика Mazaysplav');
}


// ============================================
// АВТОМАТИЧЕСКИЙ ПЕРЕСЧЕТ СУММЫ ПРИ ИЗМЕНЕНИИ
// ============================================
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  
  // Пропускаем системные листы и шапку
  const systemSheets = ['Сплавы', 'Отчеты', 'Статистика', 'Users', 'Логи', 'Архив'];
  if (systemSheets.indexOf(sheetName) !== -1 || sheetName.toLowerCase().includes('архив') || row < 2) return;
  
  // Следим за колонками: Участники (5,6), Жилье (12), Услуги (13), Цены (14,15), Оплата доставки (19)
  const trackedCols = [5, 6, 12, 13, 14, 15, 19];
  if (trackedCols.indexOf(col) === -1) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const splavSheet = ss.getSheetByName('Сплавы');
    if (!splavSheet) return;
    
    // Получаем данные строки
    const rowData = sheet.getRange(row, 1, 1, 21).getValues()[0];
    const adults = parseInt(rowData[4]) || 0;
    const children = parseInt(rowData[5]) || 0;
    const accType = String(rowData[11] || "").trim();
    const extraType = String(rowData[12] || "").trim();
    const priceA = parseFloat(rowData[13]) || 0;
    const priceC = parseFloat(rowData[14]) || 0;
    const prepayment = parseFloat(rowData[18]) || 0;
    
    // Ищем цены в листе Сплавы для этого сплава
    const splavData = splavSheet.getDataRange().getValues();
    const splavHeaders = splavData[0].map(h => String(h).trim().toLowerCase());
    
    const tripDateStr = sheetName.split(' ')[0];
    const tripTimeStr = sheetName.split(' ')[1] || "";
    
    let prices = { camper: 0, tipi: 0, safari: 0, ourTent: 0, ownTent: 0, sauna: 0 };
    let tripFound = false;
    
    for (let i = 1; i < splavData.length; i++) {
      const dVal = splavData[i][splavHeaders.indexOf('дата')];
      if (!dVal) continue;
      const dStr = Utilities.formatDate(new Date(dVal), ss.getSpreadsheetTimeZone(), "dd.MM.yyyy");
      
      let tStr = "";
      const tVal = splavData[i][splavHeaders.indexOf('время')];
      if (tVal instanceof Date) tStr = Utilities.formatDate(tVal, ss.getSpreadsheetTimeZone(), "HH:mm");
      else tStr = String(tVal).split(' ')[0];
      
      if (dStr === tripDateStr && tStr === tripTimeStr) {
        prices.camper = parseFloat(splavData[i][splavHeaders.indexOf('цена кемпера')]) || 0;
        prices.tipi = parseFloat(splavData[i][splavHeaders.indexOf('цена типи')]) || 0;
        prices.safari = parseFloat(splavData[i][splavHeaders.indexOf('цена сафари')]) || 0;
        prices.ourTent = parseFloat(splavData[i][splavHeaders.indexOf('цена вашей палатки')]) || 0;
        prices.ownTent = parseFloat(splavData[i][splavHeaders.indexOf('цена места (своя палатка)')]) || 0;
        prices.sauna = parseFloat(splavData[i][splavHeaders.indexOf('цена баня+купель')]) || 0;
        tripFound = true;
        break;
      }
    }
    
    if (!tripFound) return;
    
    // Считаем сумму жилья
    let accSum = 0;
    const accNorm = accType.toLowerCase();
    if (accNorm === 'кемпер') accSum = prices.camper;
    else if (accNorm === 'типи') accSum = prices.tipi;
    else if (accNorm === 'сафари-тент') accSum = prices.safari;
    else if (accNorm === 'наша палатка' || accNorm === 'ваша палатка') accSum = prices.ourTent;
    else if (accNorm === 'своя палатка') accSum = (adults + children) * prices.ownTent;
    
    // Считаем сумму услуг
    let extraSum = 0;
    if (extraType.toLowerCase().includes('баня')) extraSum = prices.sauna;
    
    const newTotal = (adults * priceA) + (children * priceC) + accSum + extraSum;
    const newRemaining = newTotal - prepayment;
    
    // Обновляем ИТОГО (16) и Остаток (20)
    sheet.getRange(row, 16).setValue(newTotal);
    sheet.getRange(row, 20).setValue(newRemaining);
    
    // Динамическая подсветка жилья и услуг
    const accValLC = accType.toLowerCase();
    const extraValLC = extraType.toLowerCase();
    const hasAcc = accValLC !== "" && accValLC !== "-" && accValLC !== "нет" && accValLC !== "нет жилья";
    const hasExtra = extraValLC !== "" && extraValLC !== "-" && extraValLC !== "нет";
    
    const targetRange = sheet.getRange(row, 12, 1, 2);
    if (hasAcc || hasExtra) {
      targetRange.setBackground('#dcfce7').setFontWeight('bold').setBorder(true, true, true, true, null, null, "#10b981", SpreadsheetApp.BorderStyle.SOLID);
    } else {
      targetRange.setBackground(null).setFontWeight('normal').setBorder(false, false, false, false, false, false);
    }
    
  } catch (err) {
    // В случае ошибки пишем в логи (опционально)
  }
}

// ============================================
// РЕЗЕРВНОЕ КОПИРОВАНИЕ
// ============================================

/**
 * Создает резервную копию таблицы в папку "Backups_Mazay" на Google Диске.
 */
function createBackup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss.getId();
    const file = DriveApp.getFileById(ssId);
    
    // Текущая дата для названия
    const tz = ss.getSpreadsheetTimeZone();
    const dateStr = Utilities.formatDate(new Date(), tz, "dd.MM.yyyy_HH-mm");
    const backupName = "Backup_" + ss.getName() + "_" + dateStr;
    
    // Ищем или создаем папку для бэкапов
    const folderName = "Backups_Mazay";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    // Создаем копию
    file.makeCopy(backupName, folder);
    
    logAction('BACKUP', 'Создана копия: ' + backupName, 'Успешно');
  } catch (e) {
    logAction('BACKUP_ERROR', e.toString(), 'Ошибка');
  }
}
