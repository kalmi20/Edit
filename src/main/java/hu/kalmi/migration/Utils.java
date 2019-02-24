package hu.kalmi.migration;

import java.util.Iterator;
import java.util.List;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Utils {

	private static Logger LOG = LoggerFactory.getLogger(Utils.class);

	public static boolean allCellsBlank(Row inRow, List<Cell> cells) {
		return cells.stream().filter(c -> c != null && inRow.getCell(c.getColumnIndex()) != null).allMatch(c -> {
			CellType cellType = inRow.getCell(c.getColumnIndex()).getCellType();
			return cellType.equals(CellType.BLANK);
		});
	}

	public static String getCellValueAsString(Cell cell) {
		if (cell.getCellType().equals(CellType.NUMERIC)) {
			return Integer.toString((int) Math.round(cell.getNumericCellValue()));
		} else {
			return cell.getRichStringCellValue().getString();
		}
	}

	public static Optional<String> getCellValueAsString(Row row, int index) {
		Cell cell = row.getCell(index);
		if (cell == null) {
			return Optional.empty();
		}

		if (cell.getCellType().equals(CellType.NUMERIC)) {
			return Optional.ofNullable(Integer.toString((int) Math.round(cell.getNumericCellValue())));
		} else {
			if (cell.getRichStringCellValue() == null || cell.getRichStringCellValue().getString() == null) {
				return Optional.empty();
			}
			return Optional.ofNullable(cell.getRichStringCellValue().getString());
		}
	}

	public static Cell getCellWithName(Sheet inSheet, List<String> asList) {
		for (String string : asList) {
			Cell find = getCellWithName(inSheet, string);
			if (find != null) {
				return find;
			}
		}
		return null;
	}

	public static Cell getCellWithName(Sheet sheet, String name) {
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Iterator<Cell> cellIterator = rowIterator.next().cellIterator();
			while (cellIterator.hasNext()) {
				Cell c = cellIterator.next();

				if (c.getCellType().equals(CellType.STRING)) {
					String toMatch = c.getRichStringCellValue().getString().toLowerCase().trim();
					if (toMatch.equalsIgnoreCase(name.toLowerCase().trim()))
						return c;
				}
			}
		}
		return null;
	}

	public static int getColumnIndexWithName(Sheet sheet, String name) {
		Iterator<Cell> cellIterator = sheet.getRow(0).cellIterator();
		while (cellIterator.hasNext()) {
			Cell c = cellIterator.next();

			try {
				if (c.getRichStringCellValue().getString().equalsIgnoreCase(name)) {
					return c.getColumnIndex();
				}
			} catch (Exception e) {
				LOG.error("Can not read column with index: " + c.getColumnIndex(), e);
				throw e;
			}
		}
		return -1;
	}
}
