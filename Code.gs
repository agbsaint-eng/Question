// Rudi survey -> Google Sheet
const SHEET_ID = '1-g9eHRSv4cEMFgZcmRDFf5FIpwFwvtIlJuRGp1eNdWM';
const SHEET_NAME = 'Responses';
const HEADERS = ['زمان ثبت','نام و نام خانوادگی','شهر','شماره تماس','سن','جنسیت','نوع تفریح','تفریح دیگر','زمان در هفته','روزهای هفته','شناسه'];
const AGES=['زیر ۱۸','۱۸ تا ۲۴','۲۵ تا ۳۴','۳۵ تا ۴۴','۴۵ به بالا'];
const GENDERS=['زن','مرد'];
const TIMES=['کمتر از ۲ ساعت','۲ تا ۱۰ ساعت','بیشتر از ۱۰ ساعت'];
const DAYS=['شنبه تا چهارشنبه','پنجشنبه و جمعه','هر دو'];

function clean(v, max) {
  let s = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  // جلوگیری از اجرای فرمول در شیت
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);

    // اعتبارسنجی سمت سرور
    if (AGES.indexOf(d.age) < 0 || GENDERS.indexOf(d.gender) < 0 ||
        TIMES.indexOf(d.time) < 0 || DAYS.indexOf(d.days) < 0) {
      return out('invalid');
    }
    const name = clean(d.name, 60), city = clean(d.city, 40);
    if (!name || !city) return out('invalid');
    const acts = (Array.isArray(d.acts) ? d.acts : []).slice(0, 30).map(a => clean(a, 60)).join('، ');
    if (!acts) return out('invalid');

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange('D:D').setNumberFormat('@'); // شماره تماس به‌صورت متن (صفر اول حفظ شود)
    }
    // جلوگیری از ثبت تکراری (همان شناسه)
    const id = clean(d.id, 40);
    if (id) {
      const last = Math.max(sh.getLastRow(), 1);
      const ids = sh.getRange(1, 11, last, 1).getValues().flat();
      if (ids.indexOf(id) >= 0) return out('dup');
    }
    sh.appendRow([new Date(), name, city, clean(d.phone, 14), d.age, d.gender,
                  acts, clean(d.other, 60), d.time, d.days, id]);
    return out('ok');
  } catch (err) {
    return out('error');
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// هیچ داده‌ای از طریق آدرس خوانده نمی‌شود
function doGet() { return out('Rudi survey endpoint'); }

function out(t) { return ContentService.createTextOutput(t); }
