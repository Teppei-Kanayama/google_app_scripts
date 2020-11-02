function updateTodaysTheme() {
  const spreadsheet_id = ''
  const sheet_name = 'お題'
  const theme_column = 1
  const done_flag_column = 2
  const content_start_row = 2
  const content_end_row = 10
  const display_cell_row = 1
  const display_cell_column = 10

  // get sheet object
  var spreadsheet = SpreadsheetApp.openById(spreadsheet_id);
  var sheet = spreadsheet.getSheetByName(sheet_name);

  // read the sheet
  var themes = sheet.getRange(content_start_row, theme_column, content_end_row).getValues();
  var done_flags = sheet.getRange(content_start_row, done_flag_column, content_end_row).getValues();

  next_theme = decideNextTheme(themes, done_flags)

  // update
  display_cell = sheet.getRange(display_cell_row, display_cell_column);
  display_cell.setValue(next_theme)
}

function decideNextTheme(themes, done_flags) {
  const theme_candidates = [];
  for (var row in themes) {
    theme = themes[row][0]
    done_flag = done_flags[row][0]
    if (done_flag != 1 && theme != '') {
      theme_candidates.push(theme);
    }
  }
  next_theme = theme_candidates[Math.floor(Math.random() * theme_candidates.length)]
  return next_theme
}
