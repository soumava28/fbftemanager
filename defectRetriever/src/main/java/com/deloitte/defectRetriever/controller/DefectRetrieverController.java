package com.deloitte.defectRetriever.controller;

import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RestController;

import com.deloitte.defectRetriever.model.DefectRecord;
import com.deloitte.defectRetriever.service.DefectRetrieverService;

@RestController
@RequestMapping(path = "/defectRetriever")
@CrossOrigin(origins = "*", maxAge = 3600)
public class DefectRetrieverController {

	private static final Logger LOG = Logger.getLogger(DefectRetrieverController.class.getName());

	@Autowired
	DefectRetrieverService service;

	@RequestMapping(path = "/fetchDefect")
	public List<DefectRecord> fetchDefectDump(@RequestParam String source) {
		LOG.log(Level.INFO, "Fetching defect Dump");
		List<DefectRecord> defectList = service.fetchRecords(source);
		if (defectList != null) {
			return defectList;
		}
		return null;

	}
	
	
	@RequestMapping(path = "/fetchDefectbyKey")
	public List<DefectRecord> fetchDefectDumpById(@RequestParam String defectId) {
		LOG.log(Level.INFO, "Fetching defect Dump");
		List<DefectRecord> defectList = service.fetchRecordsById(defectId);
		if (defectList != null) {
			return defectList;
		}
		return null;

	}

}
