package nchu.service;

import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import nchu.ExcelAPI;
import nchu.model.TrainingModel;

public class ReadExcelService {

	int records = 0; // 抓取的訓練資料總數
	int failedRecords = 0; // 檔案處理失敗的次數，應該要為0
	double lastSick = 0.0; // 上次的發病度
	StringBuffer text = new StringBuffer(); // feature 特徵集，每個特徵以逗號隔開，每個特徵集以換行字元隔開
	StringBuffer results = new StringBuffer(); // result 0或1的陣列，以換行字元隔開
	
	TrainingModel model; //用哪一個model去抓去特徵值，可以自行修改

	/**
	 * model有定義針對不同的version抓取對應的特徵
	 */
	public enum ModelMode {
		version1, //基本模型
		version2, //增加每日溫溼度
		version3 //original version
	}
	
	public static ModelMode mode;
	
	public ReadExcelService(ModelMode mode) {
		this.mode = mode;
		model = new TrainingModel(this);
	}

	/**
	 * 讀取檔案
	 * @param years
	 */
	public void read(String... years) {
		
		for(String year : years) {
			File file = new File(ExcelAPI.filePath + getFileNameByYear(year));
			try {
				
				Workbook wb = Workbook.getWorkbook(file);
				Sheet sheet = wb.getSheet(0);
				System.out.println("!!!: " + file);
				for(int s=0; s<wb.getSheets().length; s ++) {
					sheet = wb.getSheet(s);
					lastSick = getTotalSick(sheet.getRow(1));
					for (int i = 1; i < sheet.getRows(); i++) {
						Cell[] row = sheet.getRow(i);
						if(model.isFullData(row)) {
							doSearch(row, sheet, i);
						}
						
					}
				}
				System.out.println("資料筆數 : " + records);
				System.out.println("資料丟失筆數 : " + failedRecords);
				wb.close();
			} catch (BiffException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			
		}
		
		
	}

	/**
	 * 不同年度對應不同檔案
	 * @param year
	 * @return
	 */
	private String getFileNameByYear(String year) {
		switch(Integer.parseInt(year)) {
		case 103 :
			return "103data.xls";
		case 104 :
			return "104data.xls";
		case 105 :
			return "105data.xls";
		case 106 :
			return "106data.xls";
		case 107 :
			return "107data.xls";
		case 108 :
			return "108data.xls";
		case 109 :
			return "109data.xls";
		case 110 :
			return "110data.xls";
		default:
			return "data-0511.xls";
					
		}
	}
	
	private void doSearch(Cell[] row, Sheet sheet, int currentIndex) {
//		for(int y=0; y<row.length; y++) {
//			System.out.print(row[y].getContents() + "\t");
//		}
//		System.out.println();
		// 葉稻熱病程度
		double totalSick = getTotalSick(row);
		// 與上次的差異值
		double diff = Math.abs(totalSick - lastSick);
		if(diff > 0) {
			
			double diffPercent = diff/ Math.max(totalSick, lastSick)*100;
			// 用門檻值過濾差異不明顯的訓練資料
			if( diffPercent >= ExcelAPI.threshold ) {
				System.out.println(totalSick + "\t" + lastSick + "\t" + diff);
				System.out.println(diffPercent + "%");
				// 病情加重或減輕
				int result = totalSick > lastSick ? 1 : 0;
				
				model.appendText(result, sheet, currentIndex);
				// 保留這次的發病程度，供下筆訓練資料比較
				lastSick = totalSick;
			}
		}
	}


	/**
	 * 改為只抓取葉稻熱病欄位
	 * @param row
	 * @return
	 */
	private double getTotalSick(Cell[] row) {
		double total = 0;
		try {
//			total = Double.parseDouble(row[1].getContents())
//					+ Double.parseDouble(row[2].getContents()) + Double.parseDouble(row[3].getContents());
			// 2018/12/29 只分析 葉稻熱
			total = Double.parseDouble(row[1].getContents());
		} catch(Exception e) {
			e.printStackTrace();
		}
		return total;
	}

	public StringBuffer getText() {
		return text;
	}

	public void setText(StringBuffer text) {
		this.text = text;
	}

	public StringBuffer getResults() {
		return results;
	}

	public void setResults(StringBuffer results) {
		this.results = results;
	}

	public int getRecords() {
		return records;
	}

	public void setRecords(int records) {
		this.records = records;
	}

	public int getFailedRecords() {
		return failedRecords;
	}

	public void setFailedRecords(int failedRecords) {
		this.failedRecords = failedRecords;
	}
}
