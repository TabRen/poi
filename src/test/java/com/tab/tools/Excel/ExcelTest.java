package com.tab.tools.Excel;

import org.junit.Test;

public class ExcelTest {

  @Test
  public void readForListTest() {
    Excel excel = new Excel();
    excel.readForList("/home/tab/github/cfgdc_crm/src/main/webapp/appSupportDevice/files/全部用户.xls",
        0, true);
    excel.readForList("/Users/tab.ren/Desktop/IP规划.xls", 0, true);
    excel.readForList("/Users/tab.ren/Desktop/IP规划.xlsx", 0, true);
  }

}