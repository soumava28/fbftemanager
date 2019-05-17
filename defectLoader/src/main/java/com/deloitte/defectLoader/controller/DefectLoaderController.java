package com.deloitte.defectLoader.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.cloud.context.config.annotation.RefreshScope;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.deloitte.defectLoader.client.DefectRetrieverFeignClient;
import com.deloitte.defectLoader.model.DefectRecord;
import com.deloitte.defectLoader.service.DefectLoaderService;

@RestController
@RefreshScope
@RequestMapping(path = "/defectLoader")
@CrossOrigin(origins = "*", maxAge = 3600)
public class DefectLoaderController {

	private static final Logger LOG = Logger.getLogger(DefectLoaderController.class.getName());

	@Autowired
	DefectLoaderService service;

	@Autowired
	DefectRetrieverFeignClient defectClient;

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
			defRecord.setKey(formatter.formatCellValue(row.getCell(0)));
			defRecord.setSummary(formatter.formatCellValue(row.getCell(1)));
			defRecord.setApplicableToIE(formatter.formatCellValue(row.getCell(2)));
			defRecord.setNewJiraId(formatter.formatCellValue(row.getCell(3)));
			defRecord.setComments(formatter.formatCellValue(row.getCell(4)));
			if (!"".equals(formatter.formatCellValue(row.getCell(5)).trim())) {
				defRecord.setTrack(formatter.formatCellValue(row.getCell(5)));
			} else {
				defRecord.setTrack("NA");
			}

			defectList.add(defRecord);
		}
		if (defectList != null) {
			isDataLoaded = service.insertRecords(defectList);
		}
		return isDataLoaded;

	}

	@RequestMapping(path = "/updateDefectDump", method = RequestMethod.PUT, consumes = MediaType.APPLICATION_JSON_VALUE)
	public boolean updateDefects(@RequestBody DefectRecord records) {
		LOG.log(Level.INFO, "Updating defect dump");
		  Boolean result = false;
		  ThreadPoolExecutor executor = (ThreadPoolExecutor) Executors.newFixedThreadPool(10);
		  Callable<Boolean> callableTask = () -> {
			  return Boolean.valueOf(service.
					  updateRecords(Stream.of(records).
							  collect(Collectors.toList())));
			};
			Future<Boolean> future = executor.submit(callableTask);
			try {
			    if (future.isDone()) {
			        result = future.get();
			    }
			} catch (ExecutionException | InterruptedException e) {
			    e.printStackTrace();
			}
		return result;
	}

}
