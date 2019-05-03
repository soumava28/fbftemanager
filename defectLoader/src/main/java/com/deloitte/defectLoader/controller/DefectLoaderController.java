package com.deloitte.defectLoader.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.deloitte.defectLoader.model.DefectRecord;
import com.deloitte.defectLoader.service.DefectLoaderService;

@RestController
@RequestMapping(path = "/defectLoader")
@CrossOrigin(origins = "*", maxAge = 3600)
public class DefectLoaderController {

	private static final Logger LOG = Logger.getLogger(DefectLoaderController.class.getName());

	@Autowired
	DefectLoaderService service;

	@PostMapping(path = "/loadDefectDump")
	public boolean loadDefectDump(@RequestParam("file") MultipartFile defExcelDataFile) throws IOException {
		LOG.log(Level.INFO, "Loading defect Dump");
		boolean isDataLoaded = false;
		String fileName = defExcelDataFile.getOriginalFilename();
		System.out.println("FILE NAME ----->" + fileName);
		String[] splited = fileName.split("\\s");
		String source = splited[1];
		List<DefectRecord> defectList = new ArrayList<DefectRecord>();
		XSSFWorkbook workbook = new XSSFWorkbook(defExcelDataFile.getInputStream());
		XSSFSheet worksheet = workbook.getSheetAt(1);

		for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
			DefectRecord defRecord = new DefectRecord();
			DataFormatter formatter = new DataFormatter();
			XSSFRow row = worksheet.getRow(i);
			defRecord.setSource(source);
			defRecord.setTrack(formatter.formatCellValue(row.getCell(0)));
			defRecord.setNewJiraId(formatter.formatCellValue(row.getCell(2)));
			defRecord.setComments(formatter.formatCellValue(row.getCell(3)));
			defRecord.setSecondLevelAnalysis(formatter.formatCellValue(row.getCell(4)));
			defRecord.setDevUpdates(formatter.formatCellValue(row.getCell(5)));
			defRecord.setConfirmationRequired(formatter.formatCellValue(row.getCell(6)));
			defRecord.setWidgetName(formatter.formatCellValue(row.getCell(7)));
			defRecord.setKey(formatter.formatCellValue(row.getCell(8)));
			defRecord.setSummary(formatter.formatCellValue(row.getCell(9)));
			defRecord.setDescription(formatter.formatCellValue(row.getCell(10)));
			defRecord.setComponents(formatter.formatCellValue(row.getCell(11)));
			defRecord.setStatus(formatter.formatCellValue(row.getCell(12)));
			defRecord.setPriority(formatter.formatCellValue(row.getCell(13)));
			defRecord.setResolution(formatter.formatCellValue(row.getCell(14)));
			defRecord.setAssignee(formatter.formatCellValue(row.getCell(15)));
			defRecord.setReporter(formatter.formatCellValue(row.getCell(16)));
			defRecord.setWidgetId(formatter.formatCellValue(row.getCell(17)));
			defRecord.setDefectSeverity(formatter.formatCellValue(row.getCell(18)));
			defRecord.setRootCauseCategory(formatter.formatCellValue(row.getCell(19)));
			defRecord.setTeam(formatter.formatCellValue(row.getCell(20)));
			defRecord.setCycle(formatter.formatCellValue(row.getCell(21)));
			defRecord.setWeek(formatter.formatCellValue(row.getCell(22)));
			defRecord.setPhase(formatter.formatCellValue(row.getCell(23)));
			defRecord.setLoadDefect(formatter.formatCellValue(row.getCell(24)));
			defRecord.setSet(formatter.formatCellValue(row.getCell(25)));
			defectList.add(defRecord);
		}
		if (defectList != null) {
			isDataLoaded = service.insertRecords(defectList);
		}
		return isDataLoaded;

	}

	@RequestMapping(path = "/updateDefectDump", method = RequestMethod.POST, consumes = MediaType.APPLICATION_JSON_VALUE)
	public boolean updateDefects(@RequestBody List<DefectRecord> records) {
		LOG.log(Level.INFO, "Updating defect dump");
		boolean isDataUpdated = false;
		isDataUpdated = service.updateRecords(records);
		return isDataUpdated;

	}

}
