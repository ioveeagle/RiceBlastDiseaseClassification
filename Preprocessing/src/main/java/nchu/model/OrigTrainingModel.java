package nchu.model;

import java.util.ArrayList;
import java.util.List;
import jxl.Cell;
import jxl.Sheet;
import nchu.ExcelAPI;
import nchu.service.ReadExcelService;

/**
 * 未使用  可以把不同特徵抓取的策略存下來下來，在ReadExcelService.java的constructor決定要用哪個model
 */
public class OrigTrainingModel {

  ReadExcelService service;

  public OrigTrainingModel(ReadExcelService service) {
    this.service = service;
  }

  public void appendText(int result, Sheet sheet, int currentIndex) {
    switch(service.mode) {
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

  /**
   * original version
   * @param result
   * @param sheet
   * @param currentIndex
   */
  private void appendVersion3(int result, Sheet sheet, int currentIndex) {
    List<Cell[]> lastRows = getLastRows(sheet, currentIndex, ExcelAPI.recordDays );
    if(lastRows.size() < ExcelAPI.recordDays) {
      service.setFailedRecords(service.getFailedRecords());
      return;
    }

    // high low
    Cell[] lastOneRow = lastRows.get(0);
    Cell[] lastTwoRow = lastRows.get(1);
    Cell[] lastThreeRow = lastRows.get(2);
    Cell[] lastFourRow = lastRows.get(3);
    service.getText().append(Math.max(ExcelAPI.getHighLowAvg(lastFourRow), Math.max(ExcelAPI.getHighLowAvg(lastThreeRow),Math.max(ExcelAPI.getHighLowAvg(lastOneRow), ExcelAPI.getHighLowAvg(lastTwoRow))))).append(",");

    service.getText()
        // 自定義
        .append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫
        .append(getMinimum(lastRows, 5)).append(",") //N天內最低高溫
        .append(getMaximum(lastRows, 6)).append(",") //N天內最高低溫
        .append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫
        .append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差
        .append(getMinimum(lastRows, 7)).append(",") //N天內最低溫差
        .append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕
        .append(getMinimum(lastRows, 8)).append(",") //N天內最低高濕
        .append(getMaximum(lastRows, 9)).append(",") //N天內最高低濕
        .append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕
        .append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差
        .append(getMinimum(lastRows, 10)).append(",") //N天內最低濕差
        // 自定義大氣
        .append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫
        .append(getMinimum(lastRows, 11)).append(",") //N天內最低高溫
        .append(getMaximum(lastRows, 12)).append(",") //N天內最高低溫
        .append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫
        .append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差
        .append(getMinimum(lastRows, 13)).append(",") //N天內最低溫差
        .append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕
        .append(getMinimum(lastRows, 14)).append(",") //N天內最低高濕
        .append(getMaximum(lastRows, 15)).append(",") //N天內最高高濕
        .append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕
        .append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差
        .append(getMinimum(lastRows, 16)) //N天內最低濕差

        .append("\r\n");
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
    Cell[] lastOneRow = lastRows.get(0);

    service.getText()
        // 自定義
        .append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫 0
        .append(getMinimum(lastRows, 5)).append(",") //N天內最低高溫 1
        .append(getMaximum(lastRows, 6)).append(",") //N天內最高低溫 2
        .append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫 3
        .append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差 4
        .append(getMinimum(lastRows, 7)).append(",") //N天內最低溫差 5
        .append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕 6
        .append(getMinimum(lastRows, 8)).append(",") //N天內最低高濕 7
        .append(getMaximum(lastRows, 9)).append(",") //N天內最高低濕 8
        .append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕 9
        .append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差 10
        .append(getMinimum(lastRows, 10)).append(",") //N天內最低濕差 11
        // 自定義大氣
        .append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫 12
        .append(getMinimum(lastRows, 11)).append(",") //N天內最低高溫 13
        .append(getMaximum(lastRows, 12)).append(",") //N天內最高低溫 14
        .append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫 15
        .append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差 16
        .append(getMinimum(lastRows, 13)).append(",") //N天內最低溫差 17
        .append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕 18
        .append(getMinimum(lastRows, 14)).append(",") //N天內最低高濕 19
        .append(getMaximum(lastRows, 15)).append(",") //N天內最高高濕 20
        .append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕 21
        .append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差 22
        .append(getMinimum(lastRows, 16)) //N天內最低濕差 23

        .append("\r\n");
    service.getResults().append(result).append("\r\n");

    service.setRecords(service.getRecords());


  }

  private void appendVersion2(int result, Sheet sheet, int currentIndex) {
    List<Cell[]> lastRows = getLastRows(sheet, currentIndex, ExcelAPI.humidityDays );
    if(lastRows.size() < ExcelAPI.humidityDays) {
      service.setFailedRecords(service.getFailedRecords()+1);
      return;
    }
    // 取前一筆
    Cell[] lastOneRow = lastRows.get(0);
    Cell[] lastTwoRow = null;
    Cell[] lastThreeRow = null;

    switch(ExcelAPI.humidityDays) {
      case 1:
        // 1天 連續超過3小時 超過80
        service.getText()
            .append(lastOneRow[19].getContents()).append(",");
        break;
      case 2:
        lastTwoRow = lastRows.get(1);
        // 2天 連續超過3小時 超過80
        service.getText()
            .append(lastOneRow[19].getContents()).append(",")
            .append(lastTwoRow[19].getContents()).append(",");
        break;
      case 3:
        lastTwoRow = lastRows.get(1);
        lastThreeRow = lastRows.get(2);
        service.getText()
            .append(lastOneRow[19].getContents()).append(",")
            .append(lastTwoRow[19].getContents()).append(",")
            .append(lastThreeRow[19].getContents()).append(",");
        break;
    }

    service.getText().append(ExcelAPI.getHighLowAvg(lastOneRow)).append(",");

    service.getText()
        // 自定義
        .append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫 0
        .append(getMinimum(lastRows, 5)).append(",") //N天內最低高溫 1
        .append(getMaximum(lastRows, 6)).append(",") //N天內最高低溫 2
        .append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫 3
        .append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差 4
        .append(getMinimum(lastRows, 7)).append(",") //N天內最低溫差 5
        .append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕 6
        .append(getMinimum(lastRows, 8)).append(",") //N天內最低高濕 7
        .append(getMaximum(lastRows, 9)).append(",") //N天內最高低濕 8
        .append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕 9
        .append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差 10
//				.append(getMinimum(lastRows, 10)).append(",") //N天內最低濕差 11#
        // 自定義大氣
        .append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫 12
        .append(getMinimum(lastRows, 11)).append(",") //N天內最低高溫 13
        .append(getMaximum(lastRows, 12)).append(",") //N天內最高低溫 14
        .append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫 15
        .append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差 16
        .append(getMinimum(lastRows, 13)).append(",") //N天內最低溫差 17
        .append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕 18
//				.append(getMinimum(lastRows, 14)).append(",") //N天內最低高濕 19#
        .append(getMaximum(lastRows, 15)).append(",") //N天內最高高濕 20
        .append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕 21
        .append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差 22
        .append(getMinimum(lastRows, 16)) //N天內最低濕差 23

        .append("\r\n");
    service.getResults().append(result).append("\r\n");

    service.setRecords(service.getRecords());


  }

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
        return version1FullData(row);
      case version2:
        return version2FullData(row);
      case version3:
        return version1FullData(row);
      default:
        return false;
    }
  }

  private boolean version1FullData(Cell[] row) {
    boolean flag = false;
    String data = null;
    try {
      for(int i=0; i<ExcelAPI.columnsLimitV1; i++) {
        if(i==4 || i ==17)
          continue;
        data = row[i].getContents();
        // 107日期格式不同
        if(i != 0) {
          Double.parseDouble(data);
        }
      }
      flag = true;
    } catch(Exception e) {
//			System.out.println(data);
      flag = false;
    }
    return flag;
  }

  private boolean version2FullData(Cell[] row) {
    boolean flag = false;
    String data = null;
    try {
      for(int i=0; i<ExcelAPI.columnsLimitV2; i++) {
        if(i==4 || i ==17)
          continue;
        data = row[i].getContents();
        // 107日期格式不同
        if(i != 0) {
          Double.parseDouble(data);
        }
      }
      flag = true;
    } catch(Exception e) {
//			System.out.println(data);
      flag = false;
    }
    return flag;
  }

  private void appendTextBackup (int result, Sheet sheet, int currentIndex) {
    List<Cell[]> lastRows = getLastRows(sheet, currentIndex, ExcelAPI.recordDays );
    if(lastRows.size() < ExcelAPI.recordDays) {
      service.setFailedRecords(service.getFailedRecords());
      return;
    }
    // 取前一筆
    Cell[] lastOneRow = lastRows.get(0);
    Cell[] lastTwoRow = lastRows.get(1);
    Cell[] lastThreeRow = lastRows.get(2);
//		printRow(lastOneRow, result);
//		printRow(lastTwoRow, result);
//		printRow(lastThreeRow, result);

    service.getText()
//				.append(Double.parseDouble(lastOneRow[5].getContents())).append(",") //前一天高溫
//				.append(Double.parseDouble(lastOneRow[6].getContents())).append(",") //前一天低溫
//				.append(Double.parseDouble(lastOneRow[7].getContents())).append(",") //前一天溫差
//				.append(Double.parseDouble(lastOneRow[8].getContents())).append(",") //前一天高濕
//				.append(Double.parseDouble(lastOneRow[9].getContents())).append(",") //前一天低濕
//				.append(Double.parseDouble(lastOneRow[10].getContents())).append(",") //前一天濕差
//				.append(Double.parseDouble(lastOneRow[11].getContents())).append(",") //前一天高溫
//				.append(Double.parseDouble(lastOneRow[12].getContents())).append(",") //前一天低溫
//				.append(Double.parseDouble(lastOneRow[13].getContents())).append(",") //前一天溫差
//				.append(Double.parseDouble(lastOneRow[14].getContents())).append(",") //前一天高濕
//				.append(Double.parseDouble(lastOneRow[15].getContents())).append(",") //前一天低濕
//				.append(Double.parseDouble(lastOneRow[16].getContents())).append(",") //前一天濕差
//
//				.append(Double.parseDouble(lastTwoRow[5].getContents())).append(",") //前二天高溫
//				.append(Double.parseDouble(lastTwoRow[6].getContents())).append(",") //前二天低溫
//				.append(Double.parseDouble(lastTwoRow[7].getContents())).append(",") //前二天溫差
//				.append(Double.parseDouble(lastTwoRow[8].getContents())).append(",") //前二天高濕
//				.append(Double.parseDouble(lastTwoRow[9].getContents())).append(",") //前二天低濕
//				.append(Double.parseDouble(lastTwoRow[10].getContents())).append(",") //前二天濕差
//				.append(Double.parseDouble(lastTwoRow[11].getContents())).append(",") //前二天高溫
//				.append(Double.parseDouble(lastTwoRow[12].getContents())).append(",") //前二天低溫
//				.append(Double.parseDouble(lastTwoRow[13].getContents())).append(",") //前二天溫差
//				.append(Double.parseDouble(lastTwoRow[14].getContents())).append(",") //前二天高濕
//				.append(Double.parseDouble(lastTwoRow[15].getContents())).append(",") //前二天低濕
//				.append(Double.parseDouble(lastTwoRow[16].getContents())).append(",") //前二天濕差
//
//				.append(Double.parseDouble(lastThreeRow[5].getContents())).append(",") //前N天高溫
//				.append(Double.parseDouble(lastThreeRow[6].getContents())).append(",") //前N天低溫
//				.append(Double.parseDouble(lastThreeRow[7].getContents())).append(",") //前N天溫差
//				.append(Double.parseDouble(lastThreeRow[8].getContents())).append(",") //前N天高濕
//				.append(Double.parseDouble(lastThreeRow[9].getContents())).append(",") //前N天低濕
//				.append(Double.parseDouble(lastThreeRow[10].getContents())).append(",") //前N天濕差 7
//				.append(Double.parseDouble(lastThreeRow[11].getContents())).append(",") //前N天高溫
//				.append(Double.parseDouble(lastThreeRow[12].getContents())).append(",") //前N天低溫
//				.append(Double.parseDouble(lastThreeRow[13].getContents())).append(",") //前N天溫差
//				.append(Double.parseDouble(lastThreeRow[14].getContents())).append(",") //前N天高濕
//				.append(Double.parseDouble(lastThreeRow[15].getContents())).append(",") //前N天低濕
//				.append(Double.parseDouble(lastThreeRow[16].getContents())).append(",") //前N天濕差
        // 自定義
        .append(getMaximum(lastRows, 5)).append(",") //N天內最高高溫
        .append(getMinimum(lastRows, 5)).append(",") //N天內最低高溫
        .append(getMaximum(lastRows, 6)).append(",") //N天內最高低溫
        .append(getMinimum(lastRows, 6)).append(",") //N天內最低低溫
        .append(getMaximum(lastRows, 7)).append(",") //N天內最高溫差
        .append(getMinimum(lastRows, 7)).append(",") //N天內最低溫差
        .append(getMaximum(lastRows, 8)).append(",") //N天內最高高濕
        .append(getMinimum(lastRows, 8)).append(",") //N天內最低高濕
        .append(getMaximum(lastRows, 9)).append(",") //N天內最高低濕
        .append(getMinimum(lastRows, 9)).append(",") //N天內最低低濕
        .append(getMaximum(lastRows, 10)).append(",") //N天內最高濕差
        .append(getMinimum(lastRows, 10)).append(",") //N天內最低濕差
        // 自定義大氣
        .append(getMaximum(lastRows, 11)).append(",") //N天內最高高溫
        .append(getMinimum(lastRows, 11)).append(",") //N天內最低高溫
        .append(getMaximum(lastRows, 12)).append(",") //N天內最高低溫
        .append(getMinimum(lastRows, 12)).append(",") //N天內最低低溫
        .append(getMaximum(lastRows, 13)).append(",") //N天內最高溫差
        .append(getMinimum(lastRows, 13)).append(",") //N天內最低溫差
        .append(getMaximum(lastRows, 14)).append(",") //N天內最高高濕
        .append(getMinimum(lastRows, 14)).append(",") //N天內最低高濕
        .append(getMaximum(lastRows, 15)).append(",") //N天內最高高濕
        .append(getMinimum(lastRows, 15)).append(",") //N天內最低低濕
        .append(getMaximum(lastRows, 16)).append(",") //N天內最高濕差
        .append(getMinimum(lastRows, 16)) //N天內最低濕差

        .append("\r\n");
    service.getResults().append(result).append("\r\n");

    service.setRecords(service.getRecords()+1);


  }
}
