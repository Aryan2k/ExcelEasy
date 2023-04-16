package com.example.exceleasy;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.ServletComponentScan;

@SpringBootApplication
@ServletComponentScan

public class ExceleasyApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExceleasyApplication.class, args);
	}

}
