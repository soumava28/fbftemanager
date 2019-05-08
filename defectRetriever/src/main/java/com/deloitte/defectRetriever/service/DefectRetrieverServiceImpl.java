package com.deloitte.defectRetriever.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.deloitte.defectRetriever.dao.DefectRetrieverRepo;
import com.deloitte.defectRetriever.model.DefectRecord;

@Service
public class DefectRetrieverServiceImpl implements DefectRetrieverService {

	@Autowired
	DefectRetrieverRepo defectRecordRepo;

	@Override
	public List<DefectRecord> fetchRecords(String source) {
		List<DefectRecord> defectList = defectRecordRepo.findBySource(source);
		if (defectList != null) {
			return defectList;
		}
		return null;
	}

	@Override
	public List<DefectRecord> fetchRecordsById(String defectId) {
		List<DefectRecord> defectList = defectRecordRepo.findByKey(defectId);
		if (defectList != null) {
			return defectList;
		}
		return null;
	}

	@Override
	public List<DefectRecord> fetchAllData() {
		return defectRecordRepo.findAll();
	}

}
