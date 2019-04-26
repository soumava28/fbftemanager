package com.fteretriever.service;

import static org.hamcrest.CoreMatchers.is;

import java.time.LocalDate;
import java.util.Date;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.mongodb.core.FindAndModifyOptions;
import org.springframework.data.mongodb.core.MongoOperations;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Service;

import com.fteretriever.dao.FteRetrieverRepo;
import com.fteretriever.model.FTERecord;

@Service
public class FteRetrieverServiceImpl implements FteRetrieverService {

	@Autowired
	FteRetrieverRepo fteRetieveRepo;

	@Autowired
	private MongoOperations mongoOperation;

	@Override
	public List<FTERecord> findByWeekStDt(LocalDate weekStDt) {
		List<FTERecord> fteRecords = getDataByWeekStDt(weekStDt);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findByDateaAndtrack(LocalDate weekStDt, String track) {
		List<FTERecord> fteRecords = this.fteRetieveRepo.findByTrackAndWeekStDt(track, weekStDt);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findBytrack(String track) {
		List<FTERecord> fteRecords = this.fteRetieveRepo.findByTrack(track);
		return fteRecords;
	}

	@Override
	public List<FTERecord> getAllData() {
		List<FTERecord> fteRecords = this.fteRetieveRepo.findAll();
		return fteRecords;
	}

	@Override
	public List<FTERecord> getDataByWeekStDt(LocalDate weekStDt) {
		// Search Criteria
		Query query = new Query();
		query.addCriteria(Criteria.where("weekStDt").is(weekStDt));
		List<FTERecord> fteRecords = mongoOperation.find(query, FTERecord.class);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findByDateRange(LocalDate weekStDt, LocalDate weekEdDt) {
		Query query = new Query();
		query.addCriteria(Criteria.where("weekStDt").lte(weekEdDt).gte(weekStDt)
				.and("weekEdDt").lte(weekEdDt).gte(weekStDt));
		List<FTERecord> fteRecords = mongoOperation.find(query, FTERecord.class);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findByDateRangenTrack(LocalDate weekStDt, LocalDate weekEdDt, String track) {
		Query query = new Query();
		query.addCriteria(Criteria.where("weekStDt").lte(weekEdDt).gte(weekStDt)
				.and("weekEdDt").lte(weekEdDt).gte(weekStDt)
				.and("track").is(track));
		List<FTERecord> fteRecords = mongoOperation.find(query, FTERecord.class);
		return fteRecords;
	}

}
