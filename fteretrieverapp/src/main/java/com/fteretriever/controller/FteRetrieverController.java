package com.fteretriever.controller;

import java.text.ParseException;
import java.time.LocalDate;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.fteretriever.dao.FteRetrieverRepo;
import com.fteretriever.model.FTERecord;
import com.fteretriever.service.FteRetrieverService;

@RestController
@RequestMapping("/fetchftedata")
@CrossOrigin(origins = "*", maxAge = 3600)
public class FteRetrieverController {

	private static final Logger LOG = Logger.getLogger(FteRetrieverController.class.getName());


	
	/*@Autowired
	FteRetrieverService fteRetrieveController;*/

	@Autowired
	private FteRetrieverService fteSeriece;
	

	@RequestMapping("/byDate/")
	public List<FTERecord> getDataByDateRange(@RequestParam
			@DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
	LocalDate weekStDt) throws ParseException{
	
		List<FTERecord> fteRecords = this.fteSeriece.findByWeekStDt(weekStDt);
		return fteRecords ;
	}

	@RequestMapping("/byDatenTrack/")

	public List<FTERecord> getDataByDateRange(@RequestParam
			@DateTimeFormat(iso = DateTimeFormat.ISO.DATE)LocalDate weekStDt,
			@RequestParam String track) throws ParseException{
		List<FTERecord> fteRecords = this.fteSeriece.findByDateaAndtrack( weekStDt,track);
		return fteRecords ;
	} 

	@RequestMapping("/byTrack/")
	public List<FTERecord> getDataByDateRange(
			@RequestParam String track) throws ParseException{
		List<FTERecord> fteRecords = this.fteSeriece.findBytrack(track);
		return fteRecords ;
	} 
	
	@GetMapping("/")
	public List<FTERecord> getAllRecs() throws ParseException{
	
		List<FTERecord> fteRecords = this.fteSeriece.getAllData();
		return fteRecords ;
	}
}
