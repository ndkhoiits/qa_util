package se.niteco.qa.utils;

import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import se.niteco.qa.model.ModelObject;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * Created by khoi.nguyen on 7/6/2016.
 */
public class ExcelDeserializor<T extends ModelObject> {
	public List<T> convert(String filePath, Class<T> clazz) {
		HashMap<String, String> rowData;
		Gson gson = new Gson();
		ArrayList<T> result = new ArrayList<T>();
		List<String> headers = new ArrayList<String>();
		try {
			Workbook workbook = new XSSFWorkbook(filePath);
			Iterator<Row> iRow = workbook.getSheetAt(0).rowIterator();
			Iterator<Cell> iTopCell = iRow.next().cellIterator();
			while (iTopCell.hasNext()) {
				headers.add(iTopCell.next().getStringCellValue());
			}

			while (iRow.hasNext()) {
				Iterator<Cell> iCell = iRow.next().cellIterator();
				rowData = new HashMap<String, String>();
				while (iCell.hasNext()) {
					Cell cell = iCell.next();
					rowData.put(headers.get(cell.getColumnIndex()), getStringValue(cell));
				}

				String g = gson.toJson(rowData);
				T rowObj = gson.fromJson(g, clazz);
				result.add(rowObj);
			}

			return result;

		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	private static String getStringValue(Cell cell) {
		int type = cell.getCellType();
		switch (type) {
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			case Cell.CELL_TYPE_NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			default:
				return cell.getStringCellValue();
		}
	}
}
