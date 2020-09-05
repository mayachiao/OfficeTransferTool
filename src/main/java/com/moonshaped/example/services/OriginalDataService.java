package com.moonshaped.example.services;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.LinkedList;
import java.util.List;

import org.springframework.stereotype.Component;
import org.apache.commons.lang3.math.NumberUtils;

import com.moonshaped.example.Datas.OldData;
import com.moonshaped.example.utils.MyExcel;

@Component
public class OriginalDataService {

	// load Data

	String fileLocation = "C:\\Users\\moons\\OneDrive\\文件\\Code\\Home\\OfficeTransferTool\\src\\main\\resources\\taiwan";

	List<OldData> oldDataList = new LinkedList<OldData>();

	public void loadData() {
		System.out.println("Working Directory = " + System.getProperty("user.dir"));
		oldDataList.clear();

		final File folder = new File(fileLocation);

		listFilesForFolder(folder);
		writeDataToExcel();
		

		System.out.println("finsih" + oldDataList.size());
	}
	
	private void writeDataToExcel() {
		MyExcel excel;
		try {
			excel = new MyExcel("C:\\Users\\moons\\OneDrive\\文件\\Code\\Home\\OfficeTransferTool\\src\\main\\resources\\taiwan\\myexcel.xls");

			int rowNo = 0;
			
			for(int i = 0; i <oldDataList.size();i++)
			{
				String otherStr = "";
				String ownerStr = "";
				List<String> others = oldDataList.get(i).getOtherOwner();
				List<String> owner = oldDataList.get(i).getOwner();
				boolean isIncludeBank = false;
				for(int j= 0; j<others.size(); j++) {
					// 判斷有沒有銀行 或合作 或農會
					isIncludeBank |= isNotOnlyIncludeKeyWord(others.get(j));
					otherStr += others.get(j) +",";
				}
				
				if(!isIncludeBank)
					continue;
				
				for(int j= 0; j<owner.size(); j++) {
					ownerStr += owner.get(j) +",";
				}
				
				excel.write(new String[] { oldDataList.get(i).getBuildNo(),otherStr, ownerStr}, rowNo);// 在第1行第1個單元格寫入1,第一行第二個單元格寫入2
				rowNo++;
			}

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private boolean isNotOnlyIncludeKeyWord(String temp)
	{
		boolean bo = true;
		
		List<String> keywords = new LinkedList<String>();
		keywords.add("銀行");
		keywords.add("農會");
		keywords.add("合作");
		
		for(int i = 0; i <keywords.size();i++)
		{
			String a = keywords.get(i);
			if(temp.contains(a)) {
				bo = false;
			}
		}
		
		
		return bo;
	}

	private void listFilesForFolder(final File folder) {
		for (final File fileEntry : folder.listFiles()) {
			if (fileEntry.isDirectory()) {
				listFilesForFolder(fileEntry);
			} else {
				// System.out.println(fileEntry.getName());
				loadFileData(fileEntry);
			}
		}
	}

	private void loadFileData(final File file) {
		try {
			// FileReader fr = new FileReader(file);
			FileInputStream fis = new FileInputStream(file);
			BufferedReader br = new BufferedReader(new InputStreamReader(fis, "BIG5"));
			CreateData(br);
			/*
			 * while (br.ready()) { System.out.println(br.readLine()); }
			 */
			fis.close();

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void CreateData(BufferedReader br) {
		try {
			OldData tempObj = null;
			while (br.ready()) {
				String tempData = br.readLine();

				
				if (tempData.contains("建號")) {
					tempObj = new OldData();
					oldDataList.add(tempObj);
					tempObj.setBuildNo(tempData);
					continue;
				}
				// read owner

				if (tempData.contains("所有權部(")) {
					String counterString = tempData.replaceAll("\\D+", "");
					int ownerCounter = NumberUtils.toInt(counterString);
					List<String> ownerList = new LinkedList<String>();
					tempObj.setOwner(ownerList);
					for (int i = 0; i < ownerCounter; i++) {
						ownerList.add(br.readLine());
					}
					continue;
				}

				// read other owner

				if (tempData.contains("他項權利部(")) {
					String counterString = tempData.replaceAll("\\D+", "");
					int otherOwnerCounter = NumberUtils.toInt(counterString);
					List<String> otherOwnerList = new LinkedList<String>();
					tempObj.setOtherOwner(otherOwnerList);
					for (int i = 0; i < otherOwnerCounter; i++) {
						otherOwnerList.add(br.readLine());
					}
					continue;
				}

			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
