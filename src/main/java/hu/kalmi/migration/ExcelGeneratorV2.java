package hu.kalmi.migration;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.ApplicationArguments;
import org.springframework.stereotype.Component;

@Component
public class ExcelGeneratorV2 {
	private static Logger LOG = LoggerFactory.getLogger(ExcelGeneratorV2.class);

	private SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");

	private void copyRowContent(Row outRow, Row templateRow) {
		Iterator<Cell> cellIterator = templateRow.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = (Cell) cellIterator.next();
			Cell createCell = outRow.createCell(cell.getColumnIndex(), cell.getCellType());
			switch (cell.getCellType()) {
			case NUMERIC:
				createCell.setCellValue(cell.getNumericCellValue());
				break;
			case BLANK:
			case BOOLEAN:
			case ERROR:
			case FORMULA:
			case STRING:
			case _NONE:
			default:
				createCell.setCellValue(cell.getStringCellValue());
			}
		}
	}

	public void run(ApplicationArguments args) throws IOException, FileNotFoundException {

		List<String> folderList = args.getOptionValues("mappa");

		if (CollectionUtils.isEmpty(folderList)) {
			LOG.error("Add meg a fájlok helyét");
			System.exit(-1);
			return;
		}

		File dir = new File(folderList.get(0));
		List<File> files = Arrays.asList(dir.listFiles());

		if (files.size() == 0) {
			LOG.error("Nincsenek fajlok a mappaban");
			System.exit(-1);
			return;
		}
		List<File> inputFiles = files.stream().filter(f -> !f.getName().contains("ATMASOLT"))
				.collect(Collectors.toList());

		inputFiles.forEach(f -> {
			Workbook inWorkbook = null;
			Workbook outWorkbook = null;
			try {
				inWorkbook = WorkbookFactory.create(f);
				outWorkbook = WorkbookFactory.create(getClass().getResourceAsStream("/fill/template.xls"));

				Sheet inSheet = inWorkbook.getSheetAt(0);
				Sheet outSheet = outWorkbook.getSheetAt(0);

				List<Cell> cells = new ArrayList<Cell>();
				Cell cell1 = validateAndGetCellWIthName(inSheet, "Főkönyv");
				if (cell1 != null) {
					cells.add(cell1);
				}
				
				Cell cell2 = validateAndGetCellWIthName(inSheet, "Hrsz");
				if (cell2 != null) {
					cells.add(cell2);
				}
				
				Cell cell3 = validateAndGetCellWIthName(inSheet, "Összeg");
				if (cell3 != null) {
					cells.add(cell3);
				}
				
				Cell cell4 = validateAndGetCellWIthName(inSheet, "Ativálás esetén dátum");
				if (cell4 != null) {
					cells.add(cell4);
				}
				
				Cell cell5 = validateAndGetCellWIthName(inSheet, "Megjegyzés");
				if (cell5 != null) {
					cells.add(cell5);
				}
				
				Cell cell6 = validateAndGetCellWIthName(inSheet, "Partner név");
				if (cell6 != null) {
					cells.add(cell6);
				}
				
				Cell cell7 = validateAndGetCellWIthName(inSheet, "Költséghely");
				if (cell7 != null) {
					cells.add(cell7);
				}
				
				Cell cell8 = validateAndGetCellWIthName(inSheet, "Teljesítés dátuma");
				if (cell8 != null) {
					cells.add(cell8);
				}
				Cell cell9 = validateAndGetCellWIthName(inSheet, "Okmány száma");
				if (cell9 != null) {
					cells.add(cell9);
				}

				Cell cell10 = validateAndGetCellWIthName(inSheet, "Pályázati azonosító");
				if (cell10 != null) {
					cells.add(cell10);
				}
				Cell cell11 = validateAndGetCellWIthName(inSheet, "Gyáriszám");
				if (cell11 != null) {
					cells.add(cell11);
				}

				Cell cell12 = validateAndGetCellWIthName(inSheet, "Üzembehelyezési okmány");
				if (cell12 != null) {
					cells.add(cell12);
				}
				
				Iterator<Row> rowIterator = inSheet.rowIterator();
				// Skip trough
				while (rowIterator.hasNext() && rowIterator.next().getRowNum() < cell1.getRowIndex()) {
				}

				int rowNum = 1;
				Row templateRow = outSheet.getRow(1);
				while (rowIterator.hasNext()) {
					Row inRow = rowIterator.next();
					Row outRow = outSheet.createRow(rowNum);

					if (Utils.allCellsBlank(inRow, cells)) {
						break;
					}
					copyRowContent(outRow, templateRow);
					trySetCellValue(cell1, inRow, outRow, "L");
					trySetCellValue(cell2, inRow, outRow, "W");
					trySetCellValue(cell3, inRow, outRow, "AE");
					trySetCellValue(cell4, inRow, outRow, "AD", true);
					trySetCellValue(cell5, inRow, outRow, "J");
					trySetCellValue(cell5, inRow, outRow, "K");

					trySetCellValue(cell6, inRow, outRow, "C");
					trySetCellValue(cell7, inRow, outRow, "P");
					trySetCellValue(cell8, inRow, outRow, "AF", true);
					trySetCellValue(cell9, inRow, outRow, "BI");
					trySetCellValue(cell10, inRow, outRow, "AN");
					trySetCellValue(cell11, inRow, outRow, "T");
					trySetCellValue(cell11, inRow, outRow, "AB");

					rowNum++;
				}

				FileOutputStream fileOutStream = new FileOutputStream(
						new File(dir.getPath() + "/" + FilenameUtils.removeExtension(f.getName()) + "_ATMASOLT.xls"));
				outWorkbook.write(fileOutStream);
				fileOutStream.close();
			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} finally {
				if (inWorkbook != null) {
					try {
						inWorkbook.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
				if (outWorkbook != null) {
					try {
						outWorkbook.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
	}
	
	public static int toNumber(String name) {
        int number = 0;
        for (int i = 0; i < name.length(); i++) {
            number = number * 26 + (name.charAt(i) - ('A' - 1));
        }
        return number;
    }

	private Cell validateAndGetCellWIthName(Sheet inSheet, String name) {
		Cell cell1 = Utils.getCellWithName(inSheet, name);
		if (cell1 == null) {
			LOG.error("Can not find " + name + " column in excel");
		}
		return cell1;
	}

	private void trySetCellValue(Cell header, Row inRow, Row outRow, String columnName) {
		trySetCellValue(header, inRow, outRow, toNumber(columnName), false);
	}
	
	private void trySetCellValue(Cell header, Row inRow, Row outRow, int outCellIndex) {
		trySetCellValue(header, inRow, outRow, outCellIndex, false);
	}
	
	private void trySetCellValue(Cell header, Row inRow, Row outRow, String columnName, boolean isDate) {
		trySetCellValue(header, inRow, outRow, toNumber(columnName), isDate);
	}
	
	private void trySetCellValue(Cell header, Row inRow, Row outRow, int outCellIndex, boolean isDate) {
		if (inRow == null) {
			return;
		}
		if (header == null) {
			return;
		}
		Cell cellToCopyFrom = inRow.getCell(header.getColumnIndex());
		if (cellToCopyFrom == null) {
			return;
		}

		CellType cellType = cellToCopyFrom.getCellType();

		Cell cell = outRow.createCell(outCellIndex, cellType);
		if (cellType == null) {
			return;
		}
		switch (cellType) {
		case NUMERIC:
			if (isDate) {
				Date dateCellValue = cellToCopyFrom.getDateCellValue();
				if (dateCellValue == null) {
					return;
				}
				cell.setCellValue(format.format(dateCellValue));
			} else {
				cell.setCellValue(inRow.getCell(header.getColumnIndex()).getNumericCellValue());
			}
			break;
		case BLANK:
		case BOOLEAN:
		case ERROR:
		case FORMULA:
		case STRING:
		case _NONE:
		default:
			if (isDate) {
				try {
					Date dateCellValue = cellToCopyFrom.getDateCellValue();
					if (dateCellValue == null) {
						return;
					}
					cell.setCellValue(format.format(dateCellValue));
				} catch (Exception e) {
					cell.setCellValue(inRow.getCell(header.getColumnIndex()).getStringCellValue());
				}
			} else {
				cell.setCellValue(inRow.getCell(header.getColumnIndex()).getStringCellValue());
			}
		}

	}

}
