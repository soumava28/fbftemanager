package com.defecttoexcel.main.servicee;

import java.io.ByteArrayInputStream;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import com.defecttoexcel.main.model.DefectRecord;

@Service
public class ExcelExportServices {

	public ByteArrayInputStream customersToExcel(List<DefectRecord> defectrRecords) throws IOException {
		String[] COLUMNs = { "Key", "Track/Module", "Summary", "Applicable to IE?", "To be fixed? (IE JIRA ID)",
				"Comments", "Updated Date" };
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
			CreationHelper createHelper = workbook.getCreationHelper();

			Sheet sheet = workbook.createSheet("DefectRecord");

			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setColor(IndexedColors.BLUE.getIndex());

			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);

			// Row for Header
			Row headerRow = sheet.createRow(0);

			// Header
			for (int col = 0; col < COLUMNs.length; col++) {
				Cell cell = headerRow.createCell(col);
				cell.setCellValue(COLUMNs[col]);
				cell.setCellStyle(headerCellStyle);
			}

			// CellStyle for Age
			CellStyle ageCellStyle = workbook.createCellStyle();
			ageCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));

			int rowIdx = 1;
			for (DefectRecord defectRecord : defectrRecords) {
				Row row = sheet.createRow(rowIdx++);

				row.createCell(0).setCellValue(defectRecord.getKey());
				row.createCell(1).setCellValue(defectRecord.getTrack());
				row.createCell(2).setCellValue(defectRecord.getSummary());
				row.createCell(3).setCellValue(defectRecord.getApplicableToIE());
				row.createCell(4).setCellValue(defectRecord.getNewJiraId());
				row.createCell(5).setCellValue(defectRecord.getComments());
				if (null!=defectRecord.getUpdateDt()) {
					row.createCell(6).setCellValue(defectRecord.getUpdateDt().toString());
				}else {
					row.createCell(6).setCellValue("");
				}
				

			}

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
	}

	public ResponseEntity<InputStreamResource> saveExelData(List<DefectRecord> defectRecords) throws IOException {
		HttpHeaders headers = new HttpHeaders();
		headers.add("Content-Disposition", "attachment; filename=DefectDataDump.xlsx");

		return ResponseEntity.ok().headers(headers).body(new InputStreamResource(customersToExcel(defectRecords)));
	}
}
