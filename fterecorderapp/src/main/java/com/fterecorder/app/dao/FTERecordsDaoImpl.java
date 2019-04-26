package com.fterecorder.app.dao;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.mongodb.core.MongoOperations;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.fterecorder.app.model.FTERecord;

@Repository
public class FTERecordsDaoImpl implements FTERecordsDao {

	@Autowired
	private MongoOperations mongoOperation;
	@Override
	public void deleteRecordsById(Long id) {
		  Query query = new Query(Criteria.where("_id").is(id));
		  try {
		  mongoOperation.findAndRemove(query, FTERecord.class);
		  }catch (Exception e) {
			throw e;
		}
	}

}
