package com.defecttoexcel.main.controller;

import java.io.ByteArrayInputStream;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import com.defecttoexcel.main.model.DefectRecord;
import com.defecttoexcel.main.servicee.DataFromMicroService;
import com.defecttoexcel.main.servicee.ExcelExportServices;

@RestController
@RequestMapping(path = "/defectsToExcel")
public class DataToExcelController {

	@Autowired
	ExcelExportServices excelService;
	
	@Autowired
	DataFromMicroService dataFromServices;

	@GetMapping("/saveDataExcel")
	public ResponseEntity<InputStreamResource> dataToExcel() throws IOException {
		System.out.println("----------------------->");
		List<DefectRecord> defectRecords = dataFromServices.getTheDefectData();
		System.out.println("----------------------->" + defectRecords);
		// return IOUtils.toByteArray(in);

	return excelService.saveExelData(defectRecords);
	}

	@GetMapping("/getData")
	@ResponseBody
	public List<DefectRecord> stringReturn() throws IOException {
		System.out.println("----------------------->");
		return dataFromServices.getTheDefectData();
	}

}
