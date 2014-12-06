package com.github.sd4324530.fastexcel;

import com.github.sd4324530.fastexcel.annotation.MapperCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 快速简单操作Excel的工具
 *
 * @author peiyu
 */
public final class FastExcel {

    private static final Logger LOG = LoggerFactory.getLogger(FastExcel.class);

    /**
     * 时日类型的数据默认格式化方式
     */
    private static final DateFormat FORMAT = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

    private int startRow;

    private String sheetName;

    private final String excelFilePath;

    /**
     * 表示是否为2007以上版本的文件，因为poi处理2种文件的API不太一样
     */
    private final boolean isXlsx;

    /**
     * 构造方法，传入需要操作的excel文件路径
     *
     * @param excelFilePath 需要操作的excel文件的路径
     */
    public FastExcel(String excelFilePath) {
        this.startRow = 0;
        this.sheetName = "Sheet1";
        this.excelFilePath = excelFilePath;
        String s = this.excelFilePath.substring(this.excelFilePath.indexOf(".") + 1);
        isXlsx = s.equals("xlsx");
    }

    /**
     * 开始读取的行数，这里指的是标题内容行的行数，不是数据开始的那行
     *
     * @param startRow 开始行数
     */
    public void setStartRow(int startRow) {
        if (startRow < 1) {
            throw new RuntimeException("最小为1");
        }
        this.startRow = --startRow;
    }

    /**
     * 设置需要读取的sheet名字，不设置默认的名字是Sheet1，也就是excel默认给的名字，所以如果文件没有自已修改，这个方法也就不用调了
     *
     * @param sheetName 需要读取的Sheet名字
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * 解析读取excel文件
     *
     * @param clazz 对应的映射类型
     * @param <T>   泛型
     * @return 读取结果
     */
    public <T> List<T> parse(Class<T> clazz) {
        FileInputStream fileInputStream = null;
        POIFSFileSystem poifsFileSystem;
        Workbook workbook;
        List<T> resultList = null;
        try {
            fileInputStream = new FileInputStream(this.excelFilePath);
            if (isXlsx) {
                workbook = new XSSFWorkbook(fileInputStream);
            } else {
                poifsFileSystem = new POIFSFileSystem(fileInputStream);
                workbook = new HSSFWorkbook(poifsFileSystem);
            }
            Sheet sheet = workbook.getSheet(this.sheetName);
            if (null != sheet) {
                resultList = new ArrayList<T>(sheet.getLastRowNum() - 1);
                Row row = sheet.getRow(this.startRow);

                Map<String, Field> fieldMap = new HashMap<String, Field>();
                Map<String, String> titalMap = new HashMap<String, String>();

                Field[] fields = clazz.getDeclaredFields();
                //这里开始处理映射类型里的注解
                for (Field field : fields) {
                    if (field.isAnnotationPresent(MapperCell.class)) {
                        MapperCell mapperCell = field.getAnnotation(MapperCell.class);
                        fieldMap.put(mapperCell.cellName(), field);
                    }
                }

                for (Cell tital : row) {
                    CellReference cellRef = new CellReference(tital);
                    titalMap.put(cellRef.getCellRefParts()[2], tital.getRichStringCellValue().getString());
                }

                for (int i = this.startRow + 1; i <= sheet.getLastRowNum(); i++) {
                    T t = clazz.newInstance();
                    Row dataRow = sheet.getRow(i);
                    for (Cell data : dataRow) {
                        CellReference cellRef = new CellReference(data);
                        String cellTag = cellRef.getCellRefParts()[2];
                        String name = titalMap.get(cellTag);
                        Field field = fieldMap.get(name);
                        if (null != field) {
                            field.setAccessible(true);
                            getCellValue(data, t, field);
                        }
                    }
                    resultList.add(t);
                }
            }
        } catch (IOException e) {
            LOG.error("处理异常", e);
        } catch (InstantiationException e) {
            LOG.error("初始化异常", e);
        } catch (IllegalAccessException e) {
            LOG.error("初始化异常", e);
        } finally {
            if (null != fileInputStream) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    LOG.error("关闭流异常", e);
                }
            }
        }
        return resultList;
    }


    private void getCellValue(Cell cell, Object o, Field field) throws IllegalAccessException {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                field.set(o, cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                field.setBoolean(o, cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                field.setByte(o, cell.getErrorCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                field.set(o, cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    if (field.getType().isInstance(Date.class)) {
                        field.set(o, cell.getDateCellValue());
                    } else {
                        field.set(o, FORMAT.format(cell.getDateCellValue()));
                    }
                } else {
                    if (field.getType().isAssignableFrom(Integer.class) || field.getType().getName().equals("int")) {
                        field.setInt(o, (int) cell.getNumericCellValue());
                    } else if (field.getType().isAssignableFrom(Short.class) || field.getType().getName().equals("short")) {
                        field.setShort(o, (short) cell.getNumericCellValue());
                    } else if (field.getType().isAssignableFrom(Float.class) || field.getType().getName().equals("float")) {
                        field.setFloat(o, (float) cell.getNumericCellValue());
                    } else if (field.getType().isAssignableFrom(Byte.class) || field.getType().getName().equals("byte")) {
                        field.setByte(o, (byte) cell.getNumericCellValue());
                    } else if (field.getType().isAssignableFrom(String.class)) {
                        String s = String.valueOf(cell.getNumericCellValue());
                        if (s.contains("E")) {
                            s = s.trim();
                            BigDecimal bigDecimal = new BigDecimal(s);
                            s = bigDecimal.toPlainString();
                        }
                        field.set(o, s);
                    } else {
                        field.set(o, cell.getNumericCellValue());
                    }
                }
                break;
            case Cell.CELL_TYPE_STRING:
                field.set(o, cell.getRichStringCellValue().getString());
                break;
            default:
                field.set(o, cell.getStringCellValue());
                break;
        }
    }
}
