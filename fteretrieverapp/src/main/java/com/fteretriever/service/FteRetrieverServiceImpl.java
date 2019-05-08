package com.fteretriever.service;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.mongodb.core.MongoOperations;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Service;

import com.fteretriever.dao.FteCountRetrieverRepo;
import com.fteretriever.dao.FteRetrieverRepo;
import com.fteretriever.model.FTERecord;
import com.fteretriever.model.FteCount;

@Service
public class FteRetrieverServiceImpl implements FteRetrieverService {

	@Autowired
	FteRetrieverRepo fteRetrieveRepo;

	@Autowired
	FteCountRetrieverRepo fteCountRetrieverRepo;

	@Autowired
	private MongoOperations mongoOperation;

	@Override
	public List<FTERecord> findByWeekStDt(LocalDate weekStDt) {
		List<FTERecord> fteRecords = getDataByWeekStDt(weekStDt);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findByDateaAndtrack(LocalDate weekStDt, String track) {
		List<FTERecord> fteRecords = this.fteRetrieveRepo.findByTrackAndWeekStDt(track, weekStDt);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findBytrack(String track) {
		List<FTERecord> fteRecords = this.fteRetrieveRepo.findByTrack(track);
		return fteRecords;
	}

	@Override
	public List<FTERecord> getAllData() {
		List<FTERecord> fteRecords = this.fteRetrieveRepo.findAll();
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
		query.addCriteria(
				Criteria.where("weekStDt").lte(weekEdDt.plusDays(1)).gte(weekStDt).and("weekEdDt").lte(weekEdDt).gte(weekStDt.plusDays(1)));
		List<FTERecord> fteRecords = mongoOperation.find(query, FTERecord.class);
		return fteRecords;
	}

	@Override
	public List<FTERecord> findByDateRangenTrack(LocalDate weekStDt, LocalDate weekEdDt, String track) {
		Query query = new Query();
		query.addCriteria(Criteria.where("weekStDt").lte(weekEdDt.plusDays(1)).gte(weekStDt).and("weekEdDt").lte(weekEdDt.plusDays(1))
				.gte(weekStDt).and("track").is(track));
		List<FTERecord> fteRecords = mongoOperation.find(query, FTERecord.class);
		return fteRecords;
	}

	@Override
	public List<FteCount> findFteCountByTrack(String track) {
		List<FteCount> fteCounts = new ArrayList<>();
		fteCounts = this.fteCountRetrieverRepo.findFteCountByTrack(track);
		if (fteCounts.size() == 0) {
			FteCount f = new FteCount();
			f.setFteCount(0);
			fteCounts.add(f);
		}
		return fteCounts;
	}

}
