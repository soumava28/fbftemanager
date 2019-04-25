package com.deloitte.loginservice.controller;

import java.util.Arrays;
import java.util.logging.Level;

import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.deloitte.loginservice.model.LoginModel;

@RestController
@RequestMapping(path = "/loginservice")
@CrossOrigin(origins = "*", maxAge = 3600)
public class LoginController {

	@RequestMapping(path = "/authenticate", method = RequestMethod.POST, consumes = MediaType.APPLICATION_JSON_VALUE)
	public boolean checkLogin(@RequestBody LoginModel model) {

		boolean isSuccess = false;
		if (("user".equalsIgnoreCase(model.getUsername())) && ("Laie$@2019#".equalsIgnoreCase(model.getPassword()))) {
			isSuccess = true;
		}
		return isSuccess;
	}

}
