package nchu.model;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import jxl.Cell;
import jxl.Sheet;
import nchu.ExcelAPI;
import nchu.service.ReadExcelService;

public class TrainingModel {

	ReadExcelService service;
	
	public TrainingModel(ReadExcelService service) {
		this.service = service;
	}

	/**
	 * 針對不同version 抓取不同特徵
	 * appendVersion 會串接特徵集
	 * @param result 病情加重或減輕
	 * @param sheet
	 * @param currentIndex 這個row的index
	 */
	public void appendText(int result, Sheet sheet, int currentIndex) {
		switch(ReadExcelService.mode) {
		case version1:
			appendVersion1(result, sheet, currentIndex);
			break;
		case version2:
			appendVersion2(result, sheet, currentIndex);
			break;
		case version3:
			appendVersion3(result, sheet, currentIndex);
			break;
		default:
			
		}
	}
	
	private void appendVersion3(int result, Sheet sheet, int currentIndex) {
		List<Cell[]> lastRows = getLastRows(sheet, currentIndex, ExcelAPI.recordDays );
		if(lastRows.size() < ExcelAPI.recordDays) {
			service.setFailedRecords(service.getFailedRecords() +1);
			return;
		}
		
		// 取前一筆
		Cell[] lastOneRow = lastRows.get(0);
		Cell[] lastTwoRow = lastRows.get(1);
		Cell[] lastThreeRow = null;
		Cell[] lastFourRow = null;
		Cell[] lastFiveRow = null;
		Cell[] lastSixRow = null;
		
		service.getText()
				// 自定義
				.append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫 
				.append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫  
				.append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差
				.append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕
				.append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕
				.append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差 
				// 自定義大氣
				.append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫
				.append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫 
				.append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差  
				.append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕  
				.append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕  
				.append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差
				
				.append(getMaximum(lastRows, 17)).append(",") //N天內最高相對溼度
				
				.append(getMaximum(lastRows, 20)).append(",") //N天內最高前日溫差
				.append(getMaximum(lastRows, 21)).append(",") //N天內最高2日溫差
				;
		
//		service.text.append(getLargestNumberOfArray( ExcelAPI.getHighLowAvg(lastOneRow),ExcelAPI.getHighLowAvg(lastTwoRow) )).append(",");
				
		
		service.setText(new StringBuffer(service.getText().substring(0, service.getText().length()-1)));
		
		service.getText().append("\r\n");
		service.getResults().append(result).append("\r\n");
		
		service.setRecords(service.getRecords()+1);
	
	}
	
	private void appendVersion1(int result, Sheet sheet, int currentIndex) {
		List<Cell[]> lastRows = getLastRows(sheet, currentIndex, ExcelAPI.recordDays );
		if(lastRows.size() < ExcelAPI.recordDays) {
			service.setFailedRecords(service.getFailedRecords());
			return;
		}
		
		// 取前一筆
		Cell[] lastOneRow = lastRows.get(0); // 前一天的資料
		Cell[] lastTwoRow = lastRows.get(1); // 前兩天的資料
		Cell[] lastThreeRow = null;
		Cell[] lastFourRow = null;
		Cell[] lastFiveRow = null;
		Cell[] lastSixRow = null;
		
		service.getText()
				// 自定義
				.append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫 
				.append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫  
				.append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差
//				.append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕
				.append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕
				.append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差 
				// 自定義大氣
				.append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫
				.append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫 
				.append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差  
//				.append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕  
				.append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕  
				.append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差
				
				.append(getMaximum(lastRows, 17)).append(",") //N天內最高相對溼度
				
				.append(getMaximum(lastRows, 20)).append(",") //N天內最高前日溫差
				.append(getMaximum(lastRows, 21)).append(",") //N天內最高2日溫差
				;
		
		service.getText().append(getLargestNumberOfArray( ExcelAPI.getHighLowAvg(lastOneRow),ExcelAPI.getHighLowAvg(lastTwoRow) )).append(",");
				
		
		service.setText(new StringBuffer(service.getText().substring(0, service.getText().length()-1)));
		
		service.getText().append("\r\n");
		service.getResults().append(result).append("\r\n");
		
		service.setRecords(service.getRecords() + 1);
	
	}
	
	private void appendVersion2(int result, Sheet sheet, int currentIndex) {
		List<Cell[]> lastRows = getLastRows(sheet, currentIndex, ExcelAPI.recordDays );
		if(lastRows.size() < ExcelAPI.recordDays) {
			service.setFailedRecords(service.getFailedRecords()+1);
			return;
		}
		
		// 取前一筆
		Cell[] lastOneRow = lastRows.get(0);
		Cell[] lastTwoRow = lastRows.get(1);
		Cell[] lastThreeRow = null;
		Cell[] lastFourRow = null;
		Cell[] lastFiveRow = null;
		Cell[] lastSixRow = null;
		
		service.getText()
				// 自定義
				.append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫 
				.append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫  
				.append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差
				.append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕
				.append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕
				.append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差 
				// 自定義大氣
				.append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫
				.append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫 
				.append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差  
				.append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕  
				.append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕  
				.append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差
				
				.append(getMaximum(lastRows, 17)).append(",") //N天內最高相對溼度
				
				.append(getMaximum(lastRows, 20)).append(",") //N天內最高前日溫差
				.append(getMaximum(lastRows, 21)).append(",") //N天內最高2日溫差
				.append(getSum(lastRows, 19)).append(",") //N天內最連續濕度超過80 flag
				;
//		service.text.append(lastOneRow[19].getContents()).append(",");
//		service.text.append(lastTwoRow[19].getContents()).append(",");
		
//		service.text.append(getLargestNumberOfArray( ExcelAPI.getHighLowAvg(lastOneRow),ExcelAPI.getHighLowAvg(lastTwoRow) )).append(",");
				
		
		service.setText(new StringBuffer(service.getText().substring(0, service.getText().length()-1)));
		
		service.getText().append("\r\n");
		service.getResults().append(result).append("\r\n");
		
		service.setRecords(service.getRecords() +1);
	
	}

	/**
	 * 往前追朔幾天的資料
	 * @param sheet
	 * @param startFrom
	 * @param stopAt
	 * @return list:病情加重或減輕時，前面幾天的資料
	 */
	private List<Cell[]> getLastRows(Sheet sheet, int startFrom, int stopAt) {
		List<Cell[]> list = new ArrayList<Cell[]>();
		int s = startFrom - 1;
		int end = startFrom - 1 - stopAt;
		for(; s > end && end >= 0; s--) {
			Cell[] last = sheet.getRow(s);
			if(isFullData(last)) {
				list.add(last);
			} else {
//				break;
			}
		}
		return list;
	}
	
	/**
	 * 取最大值
	 * @param rows
	 * @param i
	 * @return
	 */
	private double getMaximum (List<Cell[]> rows, int i) {
		double max = 0.0;
		try {
			for(Cell[] row : rows) {
				max = Double.parseDouble(row[i].getContents()) > max ? Double.parseDouble(row[i].getContents()) : max;
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
		return max;
	}

	/**
	 *
	 * @param array
	 * @return Array中值最大的
	 */
	private double getLargestNumberOfArray (Double... array) {
		double max = 0.0;
		List<Double> list = new ArrayList<Double>(Arrays.asList(array));
		max = (double) Collections.max(list);
		return max;
	}
	
	/**
	 * 取最小值
	 * @param rows
	 * @param i
	 * @return
	 */
	private double getMinimum (List<Cell[]> rows, int i) {
		double min = 100;
		try {
			for(Cell[] row : rows) {
				min = Double.parseDouble(row[i].getContents()) < min ? Double.parseDouble(row[i].getContents()) : min;
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
		return min;
	}
	
	/**
	 * 取最小值
	 * @param rows
	 * @param i
	 * @return
	 */
	private double getSum (List<Cell[]> rows, int i) {
		double sum = 0;
		try {
			for(Cell[] row : rows) {
				sum += Double.parseDouble(row[i].getContents());
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
		return sum;
	}
	
	/**
	 * 取平均值
	 * @param rows
	 * @param i
	 * @return
	 */
	private double getAverage(List<Cell[]> rows, int i) {
		double avg = 0.0;
		double sum = 0.0;
		try {
			for(Cell[] row : rows) {
				sum+= Double.parseDouble(row[i].getContents());
			}
		} catch(Exception e) {
			
		}
		if(sum > 0)
			avg = sum/rows.size();
		return Math.rint(avg*100)/100;
	}
	
	
	/**
	 * check data
	 * @param row
	 * @return
	 */
	public boolean isFullData(Cell[] row) {
		switch(service.mode) {
			case version1:
			case version3:
				return version1FullData(row);
			case version2:
				return version2FullData(row);
			default:
				return false;
		}
	}

	/**
	 * 檢查是否每個需要的特徵值都有value，如果有缺值，會跳過這筆訓練資料
	 * @param row
	 * @return
	 */
	private boolean version1FullData(Cell[] row) {
		boolean flag = false;
		String data = null;
		try {
			for(int i=0; i<ExcelAPI.columnsLimitV1; i++) {
				// 這些欄位不需要有值
				if(i==0 || i==4 || i==18 || i==19)
					continue;
				if(i==17 && !ExcelAPI.isAvgHumidity)
					continue;
				data = row[i].getContents();
				Double.parseDouble(data);
			}			
			flag = true;
		} catch(Exception e) {
//			System.out.println(data);
			flag = false;
		}
		return flag;
	}

	/**
	 * 檢查是否每個需要的特徵值都有value，如果有缺值，會跳過這筆訓練資料
	 * @param row
	 * @return
	 */
	private boolean version2FullData(Cell[] row) {
		boolean flag = false;
		String data = null;
		try {
			for(int i=0; i<ExcelAPI.columnsLimitV2; i++) {
				// 這些欄位不需要有值
				if(i==0 || i==4)
					continue;
				if(i==17 && !ExcelAPI.isAvgHumidity)
					continue;
				data = row[i].getContents();
				Double.parseDouble(data);
			}			
			flag = true;
		} catch(Exception e) {
//			System.out.println(data);
			flag = false;
		}
		return flag;
	}

	/**
	 * 取得最小差異
	 * @param value1
	 * @param value2
	 * @return
	 */
	private double getMinusABS(String value1, String value2) {
		double result = 0;
		try {
			double v1 = Double.parseDouble(value1);
			double v2 = Double.parseDouble(value2);
			result = v1-v2;
		} catch(Exception e) {
			e.printStackTrace();
		}
		return Math.abs(result);
	}


}
