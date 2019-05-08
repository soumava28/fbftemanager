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

import com.defecttoexcel.main.model.Customer;
import com.defecttoexcel.main.model.DefectRecord;
@Service
public class ExcelExportServices {

	public  ByteArrayInputStream customersToExcel(List<DefectRecord> customers) throws IOException {
		String[] COLUMNs = { "Key", "Track/Module", "Summary", "Applicable to IE?" 
				, "To be fixed? (IE JIRA ID)" , "Comments" };
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
			for (DefectRecord customer : customers) {
				Row row = sheet.createRow(rowIdx++);

				row.createCell(0).setCellValue(customer.getKey());
				row.createCell(1).setCellValue(customer.getTrack());
				row.createCell(2).setCellValue(customer.getSummary());
				row.createCell(3).setCellValue(customer.getApplicableToIE());
				row.createCell(4).setCellValue(customer.getNewJiraId());
				row.createCell(5).setCellValue(customer.getComments());
				
			
			
			}

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		}
	}
	
	public ResponseEntity<InputStreamResource> saveExelData(List<DefectRecord> defectRecords) throws IOException {
		HttpHeaders headers = new HttpHeaders();
		headers.add("Content-Disposition", "attachment; filename=customers.xlsx");

		return ResponseEntity
				.ok()
				.headers(headers)
				.body(new InputStreamResource(customersToExcel(defectRecords)));
	}
}
