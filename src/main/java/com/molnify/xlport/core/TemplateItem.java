package com.molnify.xlport.core;

import java.util.HashMap;
import java.util.Map.Entry;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFTable;

/**
 * This keeps track of individual template items: 1) Named ranges, or 2) tables Allowed named ranges
 * names: "The first character of a name must be a letter or an underscore character (_). Remaining
 * characters in the name can be letters, numbers, periods, and underscore characters. In some
 * languages, Excel may replace certain characters with underscores." Source:
 * https://support.microsoft.com/en-us/office/use-names-in-formulas-9cd0e25e-88b9-46e4-956c-cc395b74582a
 *
 * @author kirsten
 */
public class TemplateItem {

  public String name, reference, sheet;
  public int startingColumn = -1;
  private final HashMap<String, int[]> tableHeaders = new HashMap<>();
  private final HashMap<String, String> tableFormulas = new HashMap<>();

  public TemplateItem(String name, String reference, String sheet) {
    if (name == null || reference == null)
      throw new IllegalArgumentException("Name and reference cannot be null");
    this.name = name;
    this.reference = reference;
    this.sheet = sheet;
  }

  public TemplateItem(XSSFTable t) {
    this.name = t.getName();
    this.reference = t.getSheetName() + "!" + t.getCellReferences().formatAsString();
    this.sheet = t.getSheetName();
    int startRow =
        t.getStartCellReference().getRow(); // , endRow = t.getEndCellReference().getRow();
    int startColumn = t.getStartCellReference().getCol(),
        endColumn = t.getEndCellReference().getCol();
    Row firstDataRow = null;
    if (t.getXSSFSheet() != null) firstDataRow = t.getXSSFSheet().getRow(startRow + 1);
    for (int i = startColumn; i <= endColumn; i++) {
      String headerName = t.getXSSFSheet().getRow(startRow).getCell(i).getStringCellValue();
      addTableHeader(headerName, startRow, i);
      if (firstDataRow != null) {
        Cell dataCell = firstDataRow.getCell(i);
        if (dataCell != null) {
          if (dataCell.getCellType() == CellType.FORMULA) {
            addTableFormula(headerName, dataCell.getCellFormula());
          }
        }
      }
    }
  }

  private void addTableHeader(String headerName, int row, int column) {
    if (row < 0 || column < 0) throw new IllegalArgumentException("Row and column needs to be > 0");
    tableHeaders.put(headerName, new int[] {row, column});
    if (startingColumn < 0) startingColumn = column;
    else startingColumn = Math.min(startingColumn, column);
  }

  private void addTableFormula(String headerName, String formula) {
    // formula = formula.replace("[]", "");
    // System.out.println("Saving formula: " + formula);
    tableFormulas.put(headerName, "=" + formula);
  }

  public Set<String> getHeaders() {
    return tableHeaders.keySet();
  }

  public String getFormulaForHeader(String headerName) {
    return tableFormulas.get(headerName);
  }

  public int[] getRowAndColumnForHeader(String headerName) {
    return tableHeaders.get(headerName);
  }

  public boolean isTable() {
    return tableHeaders.size() > 0;
  }

  // Used to determine if a table has formulas, and if so, we can wait to populate these until the
  // end
  // Addresses as table lookup bug (which is available as a test case in TestSuiteExporter)
  public boolean isTableAndHasFormulas() {
    return tableFormulas.size() > 0;
  }

  @Override
  public String toString() {
    StringBuilder buf = new StringBuilder();
    for (Entry<String, int[]> e : tableHeaders.entrySet()) {
      buf.append("<")
          .append(e.getKey())
          .append("> r: ")
          .append(e.getValue()[0])
          .append(" c:")
          .append(e.getValue()[1]);
      if (getFormulaForHeader(e.getKey()) != null)
        buf.append("---").append(getFormulaForHeader(e.getKey())).append("---");
    }
    return "TemplateItem [" + name + "] -> [" + reference + "] " + buf;
  }
}
