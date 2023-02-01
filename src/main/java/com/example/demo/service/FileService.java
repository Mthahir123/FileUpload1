package com.example.demo.service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.entity.Tutorial;
import com.example.demo.helper.CSVHelper;
import com.example.demo.helper.ExcelHelper;
import com.example.demo.repository.TutorialRepository;

@Service
public class FileService {

	@Autowired
	TutorialRepository repo;

	public void saveExcel(MultipartFile file) {
		try {
			List<Tutorial> tutorials = ExcelHelper.excelToTutorials(file.getInputStream());
			repo.saveAll(tutorials);
		} catch (IOException e) {
			throw new RuntimeException("fail to store excel data: " + e.getMessage());
		}
	}

	public void saveCSV(MultipartFile file) {
		try {
			List<Tutorial> tutorials = CSVHelper.csvToTutorials(file.getInputStream());
			repo.saveAll(tutorials);
		} catch (IOException e) {
			throw new RuntimeException("fail to store excel data: " + e.getMessage());
		}
	}

	public ByteArrayInputStream loadCSV() {
		List<Tutorial> tutorials = repo.findAll();

		ByteArrayInputStream in = CSVHelper.tutorialsToCSV(tutorials);
		return in;
	}
	
	public ByteArrayInputStream loadExcel() {
		List<Tutorial> tutorials = repo.findAll();

		ByteArrayInputStream in = ExcelHelper.tutorialsToExcel(tutorials);
		return in;
	}

	public List<Tutorial> getAllTutorials() {
		return repo.findAll();
	}
}
