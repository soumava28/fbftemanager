package com.fterecorder.app.controller;

import java.util.Arrays;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.fterecorder.app.model.FTERecord;
import com.fterecorder.app.service.FTERecorderService;

@RestController
@RequestMapping(path="/fterecorder")
@CrossOrigin(origins = "http://localhost:4200", maxAge = 3600)
public class FTERecorderController {

	@Autowired
	FTERecorderService service;
	private static final Logger LOG = Logger.getLogger(FTERecorderController.class.getName());

	@RequestMapping(path = "/", method = RequestMethod.POST, consumes = MediaType.APPLICATION_JSON_VALUE)
	public boolean insertFTERecord(@RequestBody FTERecord[] records) {
		  LOG.log(Level.INFO, "Inserting FTE"); 

		return service.insertRecords(Arrays.asList(records));

	}
	
	@GetMapping("/")
	public String test() {
		return "FteRecorderApp service running on port 8080";

	}
}
