package com.fteretriever.service;

import java.time.LocalDate;
import java.util.Date;
import java.util.List;

import com.fteretriever.model.FTERecord;

public interface FteRetrieverService {

	public abstract List<FTERecord> findByWeekStDt(LocalDate fromDate);
	public abstract List<FTERecord> findByDateaAndtrack(LocalDate fromDate,String track);
	public abstract List<FTERecord> findBytrack(String track);
	public abstract List<FTERecord> getAllData();
	public abstract List<FTERecord> getDataByWeekStDt(LocalDate weekStDt);
	public abstract List<FTERecord> findByDateRange(LocalDate weekStDt, LocalDate weekEdDt);
	public abstract List<FTERecord> findByDateRangenTrack(LocalDate weekStDt, LocalDate weekEdDt, String track);
	

}
