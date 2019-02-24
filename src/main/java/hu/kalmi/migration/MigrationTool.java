package hu.kalmi.migration;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.ApplicationArguments;
import org.springframework.stereotype.Component;

@Component
public class MigrationTool {

	private String lapszam;
	private String naplosorszam;
	private String betetlap;
	private String helyrajziSz;

	private static Logger LOG = LoggerFactory.getLogger(MigrationTool.class);

	public void run(ApplicationArguments args) throws IOException, FileNotFoundException {
		List<String> folderList = args.getOptionValues("mappa");

		if (CollectionUtils.isEmpty(folderList)) {
			LOG.error("Add meg a fájlok helyét");
			System.exit(-1);
			return;
		}

		File dir = new File(folderList.get(0));
		List<File> files = Arrays.asList(dir.listFiles()).stream().filter(f -> !f.isHidden())
				.collect(Collectors.toList());

		if (files.size() != 2) {
			LOG.error("Csak 2 fájl lehet a mappában");
			System.exit(-1);
			return;
		}

		Workbook inWorkbook = null;
		Workbook outWorkbook = null;
		File outFile = null;
		try {
			for (File file : files) {
				Workbook temp = WorkbookFactory.create(file);
				int megjegyzésIndex = Utils.getColumnIndexWithName(temp.getSheetAt(0), "Megjegyzés");
				int azonositokIndex = Utils.getColumnIndexWithName(temp.getSheetAt(0), "Azonosítók");
				if (megjegyzésIndex < 0 || azonositokIndex < 0) {
					outWorkbook = temp;
					outFile = file;
				} else {
					inWorkbook = temp;
				}
			}

			Sheet inSheet = inWorkbook.getSheetAt(0);
			Sheet outSheet = outWorkbook.getSheetAt(0);

			int megnevezesIndex = Utils.getColumnIndexWithName(inSheet, "Megnevezés");
			int helyrajziSzamIndex = Utils.getColumnIndexWithName(outSheet, "Helyrajzi szám");
			int betetlapTipusaIndex = Utils.getColumnIndexWithName(outSheet, "Betétlap típusa");

			int megjegyzésIndex = validateAndGetColumnIndexWithName(inSheet, "Megjegyzés");
			int azonositokIndex = validateAndGetColumnIndexWithName(inSheet, "Azonosítók");
			int idIndex = validateAndGetColumnIndexWithName(inSheet, "ID");
			int naplósorszámIndex = validateAndGetColumnIndexWithName(outSheet, "Naplósorszám");
			int lapszámIndex = validateAndGetColumnIndexWithName(outSheet, "Lapszám");
			int eszkozIdIndex = validateAndGetColumnIndexWithName(outSheet, "Eszköz ID");
			int ernvenyessegIdIndex = validateAndGetColumnIndexWithName(outSheet, "Érvényesség kezdete");
			int tipusIndex = validateAndGetColumnIndexWithName(outSheet, "Típus megnevezése");
			int osszerendelesTipusIndex = validateAndGetColumnIndexWithName(outSheet, "Érvényesség kezdete");

			Iterator<Row> rowIterator = inSheet.rowIterator();
			rowIterator.next();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				Cell idCell = row.getCell(idIndex);
				String id = Utils.getCellValueAsString(idCell);

				if (Strings.isBlank(id)) {
					LOG.error("Ures Id mezo a " + idCell.getRowIndex() + "-ik sorban");
					continue;
				}

				boolean success = tryInitSearchFields(megjegyzésIndex, row);
				if (!success) {
					success = tryInitSearchFieldsV2(azonositokIndex, row);
					if (!success) {
						success = tryInitSearchFieldsV3(megnevezesIndex, azonositokIndex, row);
						if (!success) {
							LOG.error("Nem talalhato a megjegyzes vagy az asonosio");
							continue;
						}
					}
				}

				boolean idFound = false;
				for (Row outRow : outSheet) {
					if (naplosorszam != null && lapszam != null) {
						Optional<String> naplósorszámValue = Utils.getCellValueAsString(outRow, naplósorszámIndex);
						if (naplósorszámValue == null || !naplósorszámValue.get().equalsIgnoreCase(naplosorszam)) {
							continue;
						}

						Optional<String> lapszámValue = Utils.getCellValueAsString(outRow, lapszámIndex);
						if (!lapszámValue.isPresent() || !lapszámValue.get().equalsIgnoreCase(lapszam)) {
							continue;
						}
					} else {
						Optional<String> betetlapTipusaValue = Utils.getCellValueAsString(outRow, betetlapTipusaIndex);
						if (betetlapTipusaValue == null || !betetlapTipusaValue.get().equalsIgnoreCase(betetlap)) {
							continue;
						}

						Optional<String> lapszámValue = Utils.getCellValueAsString(outRow, lapszámIndex);
						if (!lapszámValue.isPresent() || !lapszámValue.get().equalsIgnoreCase(lapszam)) {
							continue;
						}

						Optional<String> helyrajziSzamValue = Utils.getCellValueAsString(outRow, helyrajziSzamIndex);
						if (!helyrajziSzamValue.isPresent()
								|| !helyrajziSzamValue.get().equalsIgnoreCase(helyrajziSz)) {
							continue;
						}
					}
					idFound = true;

					if (outRow.getCell(eszkozIdIndex) == null) {
						outRow.createCell(eszkozIdIndex);
					}
					outRow.getCell(eszkozIdIndex).setCellValue(id);

					if (outRow.getCell(osszerendelesTipusIndex) == null) {
						outRow.createCell(osszerendelesTipusIndex);
					}
					outRow.getCell(osszerendelesTipusIndex).setCellValue("E");

					if (outRow.getCell(ernvenyessegIdIndex) == null) {
						outRow.createCell(ernvenyessegIdIndex);
					}
					outRow.getCell(ernvenyessegIdIndex).setCellValue("2018-01-01");

					if (outRow.getCell(tipusIndex) == null) {
						outRow.createCell(tipusIndex);
					}
					outRow.getCell(tipusIndex).setCellValue("Betétlap-Eszköz párosítás (1:null)");
				}

				if (!idFound) {
					LOG.error("Nem talalhato osszerendeles a " + lapszam + "-" + naplosorszam + " azonositohoz");
				}
			}

			FileOutputStream fileOut = new FileOutputStream(
					FilenameUtils.removeExtension(outFile.getName()) + " összerendelt" + ".xlsx");
			outWorkbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			LOG.error("Nem nyitható meg a bemeneti excel, kérlek zárd be az Excelt mielőtt futtatod!");
			throw e;
		} finally {
			if (inWorkbook != null) {
				inWorkbook.close();
			}
			if (outWorkbook != null) {
				outWorkbook.close();
			}
		}
	}

	private int validateAndGetColumnIndexWithName(Sheet inSheet, String columnName) {
		int idIndex = Utils.getColumnIndexWithName(inSheet, columnName);
		if (idIndex < 0) {
			LOG.error("Nem talalhato az " + columnName + " oszlop az excelben!");
			System.exit(-1);
		}
		return idIndex;
	}

	/**
	 * Azonositokbol olvasd ki a helyrajzi szamot, a megjegyzesbol az aposztrofok
	 * közötti a betéti lap tipus, a x.lap a lapszám
	 * 
	 * @param azonositokIndex
	 * @param row
	 * @return
	 */
	private boolean tryInitSearchFieldsV3(int megnevezesIndex, int azonositokIndex, Row row) {
		boolean success = false;
		try {
			Cell azonositoCell = row.getCell(azonositokIndex);
			RichTextString azonosito = azonositoCell.getRichStringCellValue();

			Cell megnevezesCell = row.getCell(megnevezesIndex);
			// "F" Földterület 2.lap
			RichTextString megnevezes = megnevezesCell.getRichStringCellValue();

			Pattern tipusPattern = Pattern.compile("([\"'])(?:(?=(\\\\?))\\2.)*?\\1");
			Matcher matcher = tipusPattern.matcher(megnevezes.getString());
			matcher.find();
			betetlap = matcher.group(0);

			Pattern lapszamPattern = Pattern.compile("([0-9]).lap");
			matcher = lapszamPattern.matcher(megnevezes.getString());
			matcher.find();
			lapszam = matcher.group(0).split(".")[0];

			if (azonosito == null || Strings.isBlank(azonosito.getString())) {
				LOG.error("Ures azonosito mezo a " + azonositoCell.getRowIndex() + "-ik sorban");
			}
			if (azonosito.getString().contains("Helyrajzi")) {
				Arrays.asList(azonosito.getString().split("\n")).stream().filter(s -> s.contains("Helyrajzi"))
						.forEach(s -> {
							String[] split = s.split("Helyrajzi sz.");
							helyrajziSz = split[1];
						});
				success = true;
			} else {
				success = false;
			}
		} catch (Exception e) {
			success = false;
		}
		return success;
	}

	private boolean tryInitSearchFieldsV2(int azonositokIndex, Row row) {
		boolean success = false;
		try {
			Cell azonositoCell = row.getCell(azonositokIndex);
			RichTextString azonosito = azonositoCell.getRichStringCellValue();

			if (azonosito == null || Strings.isBlank(azonosito.getString())) {
				LOG.error("Ures azonosito mezo a " + azonositoCell.getRowIndex() + "-ik sorban");
			}

			if (azonosito.getString().contains("Kataszter")) {
				Arrays.asList(azonosito.getString().split("\n")).stream().filter(s -> s.contains("Kataszter"))
						.forEach(s -> {
							String[] split = s.split("-");

							if (split.length != 3) {
								LOG.error("A " + azonositoCell.getRowIndex()
										+ "-ik sorban nem megfelelo formatumu az azonosito: " + azonosito);
							}

							naplosorszam = split[0].split("Kataszter")[1];
							betetlap = split[1];
							lapszam = split[2];
						});
				success = true;
			} else {
				success = false;
			}
		} catch (Exception e) {
			success = false;
		}
		return success;
	}

	private boolean tryInitSearchFields(int megjegyzésIndex, Row row) {
		try {
			Cell megjegyzesCell = row.getCell(megjegyzésIndex);
			RichTextString megjegyzes = megjegyzesCell.getRichStringCellValue();

			if (megjegyzes == null || Strings.isBlank(megjegyzes.getString())) {
				return false;
			}

			String[] split = megjegyzes.getString().split("-");

			if (split.length != 2) {
				LOG.error("A " + megjegyzesCell.getRowIndex() + "-ik sorban nem megfelelo formatumu a meggyejzes: "
						+ megjegyzes);
				return false;
			}

			naplosorszam = split[0];
			lapszam = split[1];

			Integer.parseInt(naplosorszam);
			Integer.parseInt(lapszam);

			return true;
		} catch (Exception e) {
			return false;
		}
	}

}
