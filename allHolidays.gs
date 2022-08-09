function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName("AllHoliday")
  sheet.getDataRange().clearContent()
  const date = new Date(new Date().getFullYear(), 0, 1, 0, 0, 0)
  const lastYear = new Date(date.getFullYear() -1, 0, 1, 0, 0, 0)
  const nextYear = new Date(date.getFullYear() + 1, 11, 31, 0, 0, 0)
  const allHolidaysArr = MakeCalender(lastYear, nextYear)
  row = allHolidaysArr.length
  column = allHolidaysArr[0].length
  sheet.getRange(1, 1, row, column).setValues(allHolidaysArr)
}

// Start = Date Object, End = Date object, Return 二重配列[[Date object],[Date Object]...]
function MakeCalender(start, end) {
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));
  //ref https://qiita.com/kznrluk/items/790f1b154d1b6d4de398
  // Start,endの日付の範囲内にある日本の祝日をピックアップして返却する関数。ソースはGoogle公式。
  const nationalHolidayEvents = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com').getEvents(new Date(start), new Date(end))
  const nationalHolidays = []
  for (const property in nationalHolidayEvents) {
    if (nationalHolidayEvents[property].getStartTime().getDay() === 6 || nationalHolidayEvents[property].getStartTime().getDay() === 0) continue
    const nationalHolidayYear = nationalHolidayEvents[property].getStartTime().getFullYear()

    const nationalHolidayMonth = nationalHolidayEvents[property].getStartTime().getMonth()
    const nationalHolidayDate = nationalHolidayEvents[property].getStartTime().getDate()
    nationalHolidays.push(new Date(nationalHolidayYear, nationalHolidayMonth, nationalHolidayDate, 0, 0, 0))

  }

  const weekend = []
  let loopDay = new Date(start)
  loopDay = new Date(loopDay.getFullYear(), loopDay.getMonth(), loopDay.getDate(), 0, 0, 0)

  while (loopDay <= new Date(end)) {
    
    if (loopDay.getDay() === 6 || loopDay.getDay() === 0){
      // new Date()で囲む事で新しい日付オブジェクトとしている = 意図しない変更が防げる
      weekend.push(new Date(loopDay))
      }
    // 上記でnew Dateで新しい日付オブジェクトとしないとここで配列の中身が変わってしまう
    loopDay = new Date((loopDay.setDate(loopDay.getDate() + 1)))
  }

  const allHolidays = [...weekend, ...nationalHolidays]
  allHolidays.sort((function (a, b) {
    // 配列の中身を日付昇順ソート
    return (a > b ? 1 : -1);
  }))
  const returnArray = transpose([allHolidays])
  
  return returnArray;
}
