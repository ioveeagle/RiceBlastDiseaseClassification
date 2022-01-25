package nchu;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import nchu.service.ReadExcelService;
import nchu.service.ReadExcelService.ModelMode;

public class ReadExcel {

	public static void main(String[] args) {
		new ReadExcel().start();

	}
	
	String outputPath = "output_data/";
	// 修改years可以決定要讀取哪些年度的資料，分別對應不同檔案
	// "107" "103","104","105","106" 第一次跑107，第二次跑103~106，表示用103~106訓練107資料
	// "105"  "103","104","106","107" 第一次跑105，第二次跑103、104、106、107，同上
	String[] years = new String[] {"103","104","105","106","107","108","109","110"};
	ModelMode mode = ModelMode.version1; // 決定用哪個model來抓取特徵，可以自己修改model內容，組合你要的特徵，這邊只是方便保留不同的特徵集
	String fileName = getNameByYear(years, false, mode); // 輸出檔案的檔名(feature)
	String fileResultName = getNameByYear(years, true, mode); // 輸出檔案的檔名(result)

	/**
	 * 程式起點
	 */
	private void start() {
		ReadExcelService service = new ReadExcelService(mode);
		service.read(years);
		
		String data = service.getText().toString().substring(0, service.getText().toString().length()-2);
		String header = getCSVHeader(data);
		
		String outputText = header + data;
//		System.out.println(outputText);
		try {
			File fout = new File(ExcelAPI.filePath + outputPath + fileName);
			FileOutputStream fos = new FileOutputStream(fout);
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
			bw.write(outputText);
			bw.close();
		} catch (FileNotFoundException e) {
			// File was not found
			e.printStackTrace();
		} catch (IOException e) {
			// Problem when writing to the file
			e.printStackTrace();
		}
		
		try {
			File fout = new File(ExcelAPI.filePath + outputPath + fileResultName);
			FileOutputStream fos = new FileOutputStream(fout);
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
			bw.write("0" + "\r\n" + service.getResults().toString());
			bw.close();
		} catch (FileNotFoundException e) {
			// File was not found
			e.printStackTrace();
		} catch (IOException e) {
			// Problem when writing to the file
			e.printStackTrace();
		}
	}
	
	private String getCSVHeader(String data) {
		int columns = data.split("\r\n")[0].split(",").length;
		String header = "";
		for(int i=0; i<columns; i++) {
			header += i + ",";
		}
		if(header.length() > 0) {
			header = header.substring(0, header.length()-1);
		}
		return header + "\r\n";
	}

	/**
	 * 建立CSV檔名
	 * @param years
	 * @param label
	 * @param mode 
	 * @return
	 */
	private String getNameByYear(String[] years, boolean label, ModelMode mode) {
		String name = "";
		String title = "";
		for(String s : years) {
			name += s.trim();
		}
		switch(mode) {
		case version1:
//			title = "a";
			title = ExcelAPI.recordDays + "_a";
			break;
		case version2:
			title = ExcelAPI.recordDays + "_";
			break;
		case version3:
			title = ExcelAPI.recordDays + "_o";
			break;
		default:
			
		}
		return title + name + ExcelAPI.threshold + (label ? "R" : "") + ".csv";
	}
}
