package com.defecttoexcel.main.servicee;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.reactive.function.client.WebClient;

import com.defecttoexcel.main.model.DefectRecord;

import reactor.core.publisher.Mono;

@Service
public class DataFromMicroService {

	WebClient.Builder builder = WebClient.builder();
	public List<DefectRecord> getTheDefectData() {
	return builder.build().get().uri("http://localhost:8090/defectfetch/defectRetriever/getAllData").retrieve().bodyToFlux(DefectRecord.class).collectList().block();
	}
}
