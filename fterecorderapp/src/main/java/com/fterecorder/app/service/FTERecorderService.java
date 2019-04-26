package com.fterecorder.app.service;

import java.util.List;

import com.fterecorder.app.model.FTERecord;

public interface FTERecorderService {
	public boolean insertRecords(List<FTERecord> records);

	public boolean deleteById(Long id);
}
