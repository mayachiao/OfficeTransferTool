package com.moonshaped.example.controllers;

import org.springframework.web.bind.annotation.RestController;

import com.moonshaped.example.services.OriginalDataService;
import com.moonshaped.example.utils.MyExcel;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

@RestController
public class HelloController {

	@Autowired
	OriginalDataService originalDataService;

	@RequestMapping("/")
	public String index() {
		return "Greetings from Spring Boot!";
	}

	@RequestMapping(value = "/getFolders", method = RequestMethod.GET)
	public String getfolders() {

		originalDataService.loadData();
		return "Greetings from Spring Boot!";
	}

	@RequestMapping(value = "/TestExcel", method = RequestMethod.GET)
	public String TestExcel() {

		MyExcel excel;
		try {
			excel = new MyExcel("C:\\Users\\moons\\OneDrive\\文件\\Code\\Home\\OfficeTransferTool\\src\\main\\resources\\taiwan\\myexcel.xls");

			excel.write(new String[] { "1", "2" }, 0);// 在第1行第1個單元格寫入1,第一行第二個單元格寫入2

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return "Greetings from Spring Boot!";
	}

}