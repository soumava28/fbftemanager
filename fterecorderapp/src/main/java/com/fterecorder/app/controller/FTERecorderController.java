package com.fterecorder.app.controller;

import java.util.Arrays;

import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.fterecorder.app.model.FTERecord;
import com.fterecorder.app.service.FTERecorderService;

@RestController
@RequestMapping(path="/fterecorder")
@CrossOrigin(origins = "*", maxAge = 3600)
public class FTERecorderController {

	@Autowired
	FTERecorderService service;
	private static final Logger LOG = Logger.getLogger(FTERecorderController.class.getName());

	@RequestMapping(path = "/", method = RequestMethod.POST, consumes = MediaType.APPLICATION_JSON_VALUE)
	public boolean insertFTERecord(@RequestBody FTERecord[] records) {
		  LOG.log(Level.INFO, "Inserting FTE"); 
		  Boolean result = false;
		  ThreadPoolExecutor executor = (ThreadPoolExecutor) Executors.newFixedThreadPool(10);
		  Callable<Boolean> callableTask = () -> {
			  return Boolean.valueOf(service.insertRecords(Arrays.asList(records)));
			};
			Future<Boolean> future = executor.submit(callableTask);
			try {
			    if (future.isDone()) {
			        result = future.get();
			        executor.shutdown();
			    }
			} catch (ExecutionException | InterruptedException e) {
			    e.printStackTrace();
			}
		return result;

	}
	
	@RequestMapping("/deleteRecord")
	public void deleteData(@RequestParam Long id) {
		this.service.deleteById(id);
	}
	
	@RequestMapping(path = "/updateRecord", method = RequestMethod.PUT, consumes = MediaType.APPLICATION_JSON_VALUE)
	public boolean updateDefects(@RequestBody FTERecord record) {
		LOG.log(Level.INFO, "Updating FTE record");
		boolean isDataUpdated = false;
		isDataUpdated = service.updateRecords(record);
		return isDataUpdated;

	}
}
