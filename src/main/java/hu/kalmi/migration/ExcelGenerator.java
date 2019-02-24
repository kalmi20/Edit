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
public class ExcelGenerator {
	private static Logger LOG = LoggerFactory.getLogger(ExcelGenerator.class);

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
				Cell cell1 = Utils.getCellWithName(inSheet, "Beruházás főkönyvi szám");
				if (cell1 == null) {
					LOG.error("Can not find Beruházás főkönyvi szám in excel");
				}
				cells.add(cell1);

				Cell cell2 = Utils.getCellWithName(inSheet, "beszerzés dátuma");
				if (cell2 == null) {
					LOG.error("Can not find beszezés dátuma in excel");
				}
				cells.add(cell2);

				Cell cell3 = Utils.getCellWithName(inSheet, "aktiválás dátuma");
				if (cell3 == null) {
					LOG.error("Can not find aktiválás dátuma in excel");
				}
				cells.add(cell3);

				Cell cell4 = Utils.getCellWithName(inSheet, "számla szerinti megnev");
				if (cell4 == null) {
					LOG.error("Can not find számla szerinti megnev in excel");
				}
				cells.add(cell4);

				Cell cell5 = Utils.getCellWithName(inSheet, "megnevezés ahogy szeretnéd");
				if (cell5 == null) {
					LOG.error("Can not find megnevezés ahogy szeretnéd in excel");
				}
				cells.add(cell5);

				Cell cell6 = Utils.getCellWithName(inSheet, "aktiválandó összeg");
				if (cell6 == null) {
					LOG.error("Can not find aktiválandó összeg in excel");
				}
				cells.add(cell6);

				Cell cell7 = Utils.getCellWithName(inSheet, "pályázati azonosító");
				if (cell7 == null) {
					LOG.error("Can not find pályázati azonosító in excel");
				}
				cells.add(cell7);

				Cell cell8 = Utils.getCellWithName(inSheet, "költséghely");
				if (cell8 == null) {
					LOG.error("Can not find költséghely in excel");
				}
				cells.add(cell8);

				Cell cell9 = Utils.getCellWithName(inSheet, Arrays.asList("számlaszám", "beszerzési számla száma"));
				if (cell9 == null) {
					LOG.error("Can not find beszerzési számla száma in excel");
				}
				cells.add(cell9);

				Cell cell10 = Utils.getCellWithName(inSheet, "Hrsz");
				if (cell10 == null) {
					LOG.error("Can not find Hrsz in excel");
				}
				cells.add(cell10);

				Cell cell11 = Utils.getCellWithName(inSheet, "gyári szám");
				if (cell11 == null) {
					LOG.error("Can not find gyári szám in excel");
				}
				cells.add(cell11);

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
					trySetCellValue(cell1, inRow, outRow, 11);
					trySetCellValue(cell2, inRow, outRow, 29, true);
					trySetCellValue(cell3, inRow, outRow, 31, true);
					trySetCellValue(cell4, inRow, outRow, 9);
					trySetCellValue(cell5, inRow, outRow, 10);
					trySetCellValue(cell6, inRow, outRow, 30);
					trySetCellValue(cell7, inRow, outRow, 39);
					trySetCellValue(cell8, inRow, outRow, 15);
					trySetCellValue(cell9, inRow, outRow, 60);
					trySetCellValue(cell10, inRow, outRow, 22);
					trySetCellValue(cell11, inRow, outRow, 19);
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

	private void trySetCellValue(Cell header, Row inRow, Row outRow, int outCellIndex) {
		trySetCellValue(header, inRow, outRow, outCellIndex, false);
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
