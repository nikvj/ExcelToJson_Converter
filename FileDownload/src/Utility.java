import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import java.nio.file.Files;
import java.nio.file.Paths;

import com.github.opendevl.JFlat;

public class Utility {

	public static void main(String[] args) throws Exception {

		Utility json2csv = new Utility();
		String str = new String(Files.readAllBytes(Paths.get("Sensor.json")));
		json2csv.JsonToExcel(str);
		ExcelToJson();
	}

//	ExcelToJson Converter
	@SuppressWarnings("unchecked")
	public static JSONObject ExcelToJson() {
		JSONObject sheetsJsonObject = new JSONObject();
		XSSFWorkbook workbook = null;
		String sourceFilePath = "Sensor.xlsx";
		FileInputStream fileInputStream = null;

		try {
			fileInputStream = new FileInputStream(sourceFilePath);
			workbook = new XSSFWorkbook(fileInputStream);

		} catch (IOException e) {
			e.printStackTrace();
		}

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			JSONArray sheetArray = new JSONArray();
			ArrayList<String> columnNames = new ArrayList<String>();
			Sheet sheet = workbook.getSheetAt(i);
			Iterator<Row> sheetIterator = sheet.iterator();

			while (sheetIterator.hasNext()) {
				Row currentRow = sheetIterator.next();
				JSONObject jsonObject = new JSONObject();

				if (currentRow.getRowNum() != 0) {
					for (int j = 0; j < columnNames.size(); j++) {

						if (currentRow.getCell(j) != null) {
							if (currentRow.getCell(j).getCellType() == CellType.STRING) {
								jsonObject.put(columnNames.get(j), currentRow.getCell(j).getStringCellValue());
							} else if (currentRow.getCell(j).getCellType() == CellType.NUMERIC) {
								jsonObject.put(columnNames.get(j), currentRow.getCell(j).getNumericCellValue());
							} else if (currentRow.getCell(j).getCellType() == CellType.BOOLEAN) {
								jsonObject.put(columnNames.get(j), currentRow.getCell(j).getBooleanCellValue());
							} else if (currentRow.getCell(j).getCellType() == CellType.BLANK) {
								jsonObject.put(columnNames.get(j), "");
							}
						} else {
							jsonObject.put(columnNames.get(j), "");
						}
					}
					sheetArray.add(jsonObject);
				} else {
					// store column names
					for (int k = 0; k < currentRow.getPhysicalNumberOfCells(); k++) {
						columnNames.add(currentRow.getCell(k).getStringCellValue());
					}
				}
			}
			sheetsJsonObject.put(workbook.getSheetName(i), sheetArray);
		}
		try (FileWriter file = new FileWriter("Sensor.json")) {

			file.write(sheetsJsonObject.toJSONString());
			file.flush();

		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("Excel to Json success!!");
		return sheetsJsonObject;
	}

//	JsonToExcel Converter
	public void JsonToExcel(String str) throws Exception {

		JFlat flatMe = new JFlat(str);

		flatMe.json2Sheet().headerSeparator("_").getJsonAsSheet();

		flatMe.write2csv("test2.csv");

		System.out.println("Json to Excel success!!");
	}

}
