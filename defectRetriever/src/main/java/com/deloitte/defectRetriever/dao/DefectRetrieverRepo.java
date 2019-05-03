package com.deloitte.defectRetriever.dao;

import java.util.List;

import org.springframework.data.mongodb.repository.MongoRepository;

import com.deloitte.defectRetriever.model.DefectRecord;

public interface DefectRetrieverRepo extends MongoRepository<DefectRecord, String> {
	List<DefectRecord> findBySource(String source);

	List<DefectRecord> findByKey(String defectId);
}
