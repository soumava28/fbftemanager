package com.defecttoexcel.main;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@ComponentScan("com.defecttoexcel.main.*")
public class DatatoexcelApplication {

	public static void main(String[] args) {
		SpringApplication.run(DatatoexcelApplication.class, args);
	}

}
