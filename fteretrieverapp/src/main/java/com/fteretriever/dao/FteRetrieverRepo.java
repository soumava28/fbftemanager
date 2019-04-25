package com.fteretriever.dao;

import java.time.LocalDate;
import java.util.Date;
import java.util.List;

import org.springframework.data.mongodb.repository.MongoRepository;
import org.springframework.stereotype.Repository;

import com.fteretriever.model.FTERecord;

@Repository
public interface FteRetrieverRepo extends MongoRepository<FTERecord,Long> {
	public List<FTERecord> findByWeekStDt(LocalDate weekStDt);
	public List<FTERecord> findByTrack(String track);
	public List<FTERecord> findByTrackAndWeekStDt(String track, LocalDate weekStDt);
	public List<FTERecord> findByTrackOrWeekStDt(String track, LocalDate weekStDt);
}
