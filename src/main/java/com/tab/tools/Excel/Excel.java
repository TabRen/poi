package com.tab.tools.Excel;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel
 * Created by tab.ren on 2018/7/17.
 */
public class Excel {

  private final static Logger logger = LoggerFactory.getLogger(Excel.class);

  //Read row for map key: row index
  public List<Map<String, String>> readForList(String path, int sheetIndex, boolean hasHead) {
    List<Map<String, String>> readForList;
    //判断文件的扩展名
    String extName = path;
    extName = extName.substring(path.lastIndexOf("."));
    switch (extName) {
      case ".xls":
        logger.info("Excel readForList xls format detect");
        readForList = readXlsForList(path, sheetIndex, hasHead);
        break;
      case ".xlsx":
        logger.info("Excel readForList xlsx format detect");
        readForList = readXlsxForList(path, sheetIndex, hasHead);
        break;
      default:
        logger.error("Excel readForList can not detect extName");
        return null;
    }
    return readForList;
  }

  //readXlsForMap
  private List<Map<String, String>> readXlsForList(String path, int sheetIndex, boolean hasHead) {
    List<Map<String, String>> readXlsForList = new ArrayList<>();
    Map<Integer, String> head = new HashMap<>();
    try {
      FileInputStream fileInputStream = new FileInputStream(path);
      HSSFWorkbook excel = new HSSFWorkbook(fileInputStream);
      //获取sheet页
      HSSFSheet sheet = excel.getSheetAt(sheetIndex);
      for (Row aSheet : sheet) {
        HSSFRow row = (HSSFRow) aSheet;
        Iterator iterator = row.cellIterator();
        Map<String, String> map = new HashMap<>();
        while (iterator.hasNext()) {
          HSSFCell cell = (HSSFCell) iterator.next();
          if (hasHead) {
            //有表头的是情况
            if (row.getRowNum() == 0) {
              head.put(cell.getColumnIndex(), cell.getStringCellValue());
            } else {
              map.put(head.get(cell.getColumnIndex()), cell.getStringCellValue());
            }

          } else {
            //没有表头的情况
            map.put(row.getRowNum() + "-" + cell.getColumnIndex(), cell.getStringCellValue());
          }
        }
        if (hasHead && (row.getRowNum() == 0)) {
          continue;
        }
        readXlsForList.add(map);
      }
    } catch (Exception e) {
      readXlsForList.clear();
      logger.error("Excel readXlsForMap occur exception: {}", e);
    }
    return readXlsForList;
  }

  //readXlsxForMap
  private List<Map<String, String>> readXlsxForList(String path, int sheetIndex, boolean hasHead) {
    List<Map<String, String>> readXlsxForList = new ArrayList<>();
    Map<Integer, String> head = new HashMap<>();
    try {
      OPCPackage pkg = OPCPackage.open(path);
      XSSFWorkbook excel = new XSSFWorkbook(pkg);
      //获取sheet
      XSSFSheet sheet = excel.getSheetAt(sheetIndex);
      for (Row aSheet : sheet) {
        XSSFRow row = (XSSFRow) aSheet;
        Iterator iterator = row.cellIterator();
        Map<String, String> map = new HashMap<>();
        while (iterator.hasNext()) {
          XSSFCell cell = (XSSFCell) iterator.next();
          if (hasHead) {
            //有表头的是情况
            if (row.getRowNum() == 0) {
              head.put(cell.getColumnIndex(), cell.getStringCellValue());
            } else {
              map.put(head.get(cell.getColumnIndex()), cell.getStringCellValue());
            }

          } else {
            //没有表头的情况
            map.put(row.getRowNum() + "-" + cell.getColumnIndex(), cell.getStringCellValue());
          }
        }
        if (hasHead && (row.getRowNum() == 0)) {
          continue;
        }
        readXlsxForList.add(map);
      }
    } catch (Exception e) {
      logger.error("Excel readXlsxForList occur exception: {}", e);
    }
    return readXlsxForList;
  }
}
