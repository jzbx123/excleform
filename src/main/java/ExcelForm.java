import com.sun.rowset.internal.Row;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelForm {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    public static void main(String[] args) throws BiffException, IOException {
        int[] opt1={1,194};
        int[] opt2={1,22};
        MergeExcel("C:\\Users\\DELL\\Desktop\\a1.xlsx","C:\\Users\\DELL\\Desktop\\a2.xlsx","C:\\Users\\DELL\\Desktop\\a3.xlsx",
                opt1,opt2);
    }

    /*
     * 读取主键然后在控制台输出 ‘’，格式的
     */
    public static String getSheet(String sh) throws BiffException, IOException {
        String[] result;
        //用的时候注意路径
        Workbook book = Workbook.getWorkbook(new File("C:\\Users\\DELL\\Desktop\\test.xls"));
        //获得excel文件的sheet表
        Sheet sheet = book.getSheet(sh);
        int rows = sheet.getRows();//行
        System.out.println("总行数:" + rows);
        System.out.println("----------------------------");
        result = new String[rows];
        int i = 0;
        //循环读取数据
        for (i = 0; i < rows; i++) {
            //getCell(x,y)   第y行的第x列
            result[i] = new String(sheet.getCell(0, i).getContents());
        }

        StringBuffer result1 = new StringBuffer();
        for (String va : result) {
            result1.append("'" + va + "',");
        }
        return result1.toString();
    }

    /**
     * 读取xmls文件
     */
    public static List<String[]> getExcelContext(String filepath,int[] opt) {

        FileInputStream in = null;
        XSSFWorkbook workbook = null;
        try {
            in = new FileInputStream(filepath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook = new XSSFWorkbook(in);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
        List<String[]> list = new ArrayList<String[]>();
        if (workbook != null) {
            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                //获得当前sheet工作表

                XSSFSheet sheet = workbook.getSheetAt(sheetNum);
                if (sheet == null) {
                    continue;
                }
                //获得当前sheet的开始行
                int firstRowNum = opt[0]-1;
                //获得当前sheet的结束行
                int lastRowNum  = opt[1]-1;
                //循环所有行
                for (int rowNum = firstRowNum ; rowNum <= lastRowNum; rowNum++) { 
                    //获得当前行
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    //获得当前行的开始列
                    int firstCellNum = row.getFirstCellNum();
                    //获得当前行的列数
                    int lastCellNum = row.getLastCellNum();//为空列获取
//                    int lastCellNum = row.getPhysicalNumberOfCells();//为空列不获取
//                    String[] cells = new String[row.getPhysicalNumberOfCells()];
                    String[] cells = new String[row.getLastCellNum()];
                    //循环当前行
                    for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
                        Cell cell = row.getCell(cellNum);
                        cells[cellNum] = getCellValue(cell);
                    }
                    list.add(cells);
                }
            }
        }

        return list;
    }

    // 判断单元格类型转成字符串
    public static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }
        //判断数据的类型
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC: //数字
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING: //字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN: //Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA: //公式
//                cellValue = String.valueOf(cell.getCellFormula());
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BLANK: //空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: //故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }

    public static void MergeExcel(String mainExcel, String otherExcel,String topath,int[] opt1 ,int [] opt2 ) {

        //获取两个表数据
        List<String[]> mainlist=new ArrayList<String[]>();
        List<String[]> otherlist=new ArrayList<String[]>();
        mainlist=getExcelContext(mainExcel,opt1);
        otherlist=getExcelContext(otherExcel,opt2);
        //创建excel文件
        File NewxlsFile = new File(topath);
        // 创建一个工作簿
        XSSFWorkbook Newworkbook = new XSSFWorkbook();
        // 创建一个工作表
        XSSFSheet Newsheet = Newworkbook.createSheet("sheet1");
         //合并表
           int row=1;//记录行数
            int mainsize=mainlist.size();
            int othersize= otherlist.size();
            for (int i=0;i<mainsize;i++){

                for(int j=0;j<othersize;j++){

                    //创建行
                    XSSFRow Newrows = Newsheet.createRow(row);
                    int m=0;
                    for( m=0;m<mainlist.get(i).length;m++){
                        String va=mainlist.get(i)[m];
                        Newrows.createCell(m).setCellValue(va);
                    }
                    int next=m;
                    for(int n=0;n<otherlist.get(j).length;n++){
                        String va=otherlist.get(j)[n];
                        Newrows.createCell(next).setCellValue(va);
                        next++;
                    }
                    row++;

                }
            }
            //将excel写入
            FileOutputStream fileOutputStream = null;
            try {
                fileOutputStream = new FileOutputStream(NewxlsFile);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            try {
                Newworkbook.write(fileOutputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }

    }
}





