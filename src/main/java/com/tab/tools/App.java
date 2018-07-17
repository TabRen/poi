package com.tab.tools;

import com.tab.tools.Excel.Excel;

/**
 * com.tab.tools.App
 */
public class App {

  public static void main(String[] args) {
    Excel excel = new Excel();
    excel.readForList("/Users/tab.ren/Desktop/IP规划.xls", 0, true);
    excel.readForList("/Users/tab.ren/Desktop/IP规划.xlsx", 0, true);
  }
}
