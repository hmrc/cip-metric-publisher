function populateTable(table, values) {
  for (var r = 0; r < values.length; r++) {
    // Extend the slide table if it's too short
    if (table.getNumRows() < (r+1)) {
      table.appendRow()
    }

    row = table.getRow(r)
   
    for (var c = 0; c < table.getNumColumns(); c++) {
      if (c < values[r].length) {
        row.getCell(c).getText().setText(values[r][c])
      } else {
        row.getCell(c).getText().setText("")
      }
    }
  }

  // Delete extra rows in table
  while (table.getNumRows()>r) {
    table.getRow(table.getNumRows()-1).remove()
  }
}