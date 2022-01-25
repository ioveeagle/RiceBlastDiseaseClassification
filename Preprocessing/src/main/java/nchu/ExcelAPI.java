package nchu;

import jxl.Cell;

/**
 * 維護參數
 * 大部分API都搬到model去了
 */
public class ExcelAPI {

	static public final String filePath = "C:/Users/KYU_NLPLAB/Desktop/RiceBlastDiseaseClassification/data/";
	static public final int columnsLimitV1 = 22; //excel最後一個欄位的index，不同版本的model欄位數不同
	static public final int columnsLimitV2 = 22; //版本2的excel最後一個欄位index
	
	static public final int recordDays = 2; //取得幾天內的資料
	static public final int humidityDays = 0; // 當值等於3的時候表示：3天內有連續超過3小時 濕度超過80的情況發生
	static public final int threshold = 15; // 發病度與上一次差異的門檻值，差異小getFileNameByYear於這個值的資料不會被存起來當訓練資料
	static public final boolean isAvgHumidity = true; // 是否要把平均濕度當作特徵
	
	/**
	 * (最高溫 - 最低溫) / 高溫 * 平均溫度
	 * @param row
	 * @return
	 */
	static public double getHighLowAvg(Cell row[]) {
		double value = 0.0;
		try {
			double highTemp = Double.parseDouble(row[5].getContents());
			double lowTemp = Double.parseDouble(row[6].getContents());
			value = Math.round(( highTemp - lowTemp ) / highTemp *100 *100)/100.0;
		} catch(Exception e) {
			e.printStackTrace();
		}
		return value;
	}

}
