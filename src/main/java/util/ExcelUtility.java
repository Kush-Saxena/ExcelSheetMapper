package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import constants.InitParameters;

public class ExcelUtility {

	private String inputFilePath = InitParameters.inpFilePath;
	private String outputFilePath = InitParameters.outFilePath;
	private String sheetName = InitParameters.sheetName;

	private Workbook outputWorkBook = null;
	private Workbook inputWorkBook = null;

	// list of columns to map
	private Map<String, String> headerRowColumnNameMapping = InitParameters.colNamesMap;

	private Map<String, Integer> headerRowColumnIndexMapping = null;
	// used for mapping the inp to output cols.

	private Map<String, String> scenarioNames = null;

	private Sheet outputSheet = null;

	private int scenarioRow = InitParameters.scenarioRow;
	private int scenarioCol = InitParameters.scenarioCol;

	public ExcelUtility(boolean isXssf) throws IOException {
		loadInputWorkBook(isXssf);
		initializeOutputWorkbook(isXssf);
	}

	public void mapExcelData() throws IOException {
		mapWorkbook();
		FileOutputStream outstream = new FileOutputStream(outputFilePath);
		outputWorkBook.write(outstream);
		outputWorkBook.close();
	}

	private void initializeOutputWorkbook(boolean xssf) {
		if (xssf) {
			outputWorkBook = new HSSFWorkbook(); // XSSF Workbook
		} else {
			outputWorkBook = new HSSFWorkbook();
		}
		scenarioNames = new HashMap<String, String>();
		Iterator<String> scNumbers = InitParameters.scenarioNumbers.iterator();
		while (scNumbers.hasNext()) {
			String currSc = scNumbers.next();
			scenarioNames.put(currSc, currSc);
		}
		outputSheet = outputWorkBook.createSheet();
	}

	private void loadInputWorkBook(boolean xssf) throws IOException {
		File inputFile = new File(inputFilePath);
		FileInputStream inpStream = new FileInputStream(inputFile);
		if (xssf) {
			inputWorkBook = new HSSFWorkbook(inpStream); // new XSFFWorkbook(inpStream);
		} else {
			inputWorkBook = new HSSFWorkbook(inpStream);
		}
	}

	private void mapWorkbook() {
		// read sheet
		Sheet sh = inputWorkBook.getSheet(sheetName);
		Row headerRow = sh.getRow(scenarioRow);

		// First mapping header rows for column names.
		Iterator<Cell> headerCells = headerRow.cellIterator();
		while (headerCells.hasNext()) {
			Cell currCell = headerCells.next();
			String content = readCellContent(currCell);
			mapHeader(content, currCell.getColumnIndex());
		}

		// Next iterating the rows.
		Iterator<Row> rows = sh.rowIterator();

		while (rows.hasNext()) {
			Row row = rows.next();
			// Skip rows before the header number starts
			if (row.getRowNum() <= headerRow.getRowNum()) {
				continue;
			}

			String cellContent = readCellContent(row.getCell(scenarioCol));
			if (scenarioNames.containsKey(cellContent)) {
				// process that specific row
				Double flowOrder = Double.parseDouble(scenarioNames.get(cellContent));
				if ((flowOrder - flowOrder.intValue()) == 0.0d) {
					processRow(row, String.format("%.0f", flowOrder));
				} else {
					processRow(row, String.format("%.1f", flowOrder));
				}
				flowOrder = flowOrder + 0.1d;
				scenarioNames.put(cellContent, String.format("%.1f", flowOrder));
			}
		}
		Sheet outputSheet = outputWorkBook.getSheetAt(0);
		outputSheet.removeRow(outputSheet.getRow(0));
		outputSheet.shiftRows(1, outputSheet.getLastRowNum(), -1);
	}

	private void mapHeader(String content, Integer index) {
		if (headerRowColumnNameMapping.containsKey(content)) {
			Row outputHeaderRow = outputSheet.getRow(0);
			if (outputHeaderRow == null) {
				outputHeaderRow = outputSheet.createRow(0);
			}
			Row inputHeaderRow = outputSheet.getRow(1);
			if (inputHeaderRow == null) {
				inputHeaderRow = outputSheet.createRow(1);
			}
			int indexToPutCell = outputHeaderRow.getLastCellNum() == -1 ? outputHeaderRow.getLastCellNum() + 1
					: outputHeaderRow.getLastCellNum();
			Cell currCell = outputHeaderRow.createCell(indexToPutCell);
			currCell.setCellValue(content);
			Cell currOutputCell = inputHeaderRow.createCell(indexToPutCell);
			currOutputCell.setCellValue(headerRowColumnNameMapping.get(content));
			if (headerRowColumnIndexMapping == null) {
				headerRowColumnIndexMapping = new HashMap<String, Integer>();
			}
			headerRowColumnIndexMapping.put(content, index);
		}
	}

	private void processRow(Row row, String flowOrder) {
		Row r = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
		// set Flow order
		r.createCell(0).setCellValue(flowOrder);
		Row outputHeaderRow = outputSheet.getRow(0);
		for (int i = 1; i < outputHeaderRow.getLastCellNum(); i++) {
			Cell currOutputCell = r.createCell(i);
			String cellContent = readCellContent(outputHeaderRow.getCell(i));
			int cellIndex = headerRowColumnIndexMapping.get(cellContent);
			String cellDataToPut = readCellContent(row.getCell(cellIndex));
			currOutputCell.setCellValue(cellDataToPut);
		}
	}

	private String readCellContent(Cell currCell) {
		String content = null;
		if (currCell == null) {
			return null;
		}
		if (currCell.getCellType() == CellType.NUMERIC) {
			content = new Double(currCell.getNumericCellValue()).toString();
		} else if (currCell.getCellType() == CellType.ERROR) {
			// Skip
		} else {
			content = currCell.getRichStringCellValue().getString();
		}
		return content;
	}
}
