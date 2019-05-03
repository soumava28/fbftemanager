package com.deloitte.defectRetriever.service;

import java.util.List;

import com.deloitte.defectRetriever.model.DefectRecord;

public interface DefectRetrieverService {

	public List<DefectRecord> fetchRecords(String source);

	public List<DefectRecord> fetchRecordsById(String defectId);

}
