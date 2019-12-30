

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelUtils {
	
    private static HSSFWorkbook workbook = null;  
    
    /** 
     * 判断文件是否存在. 
     * @param fileDir  文件路径 
     * @return 
     */  
    public static boolean fileExist(String fileDir){  
         boolean flag = false;  
         File file = new File(fileDir);  
         flag = file.exists();  
         return flag;  
    }  
    /** 
     * 判断文件的sheet是否存在. 
     * @param fileDir   文件路径 
     * @param sheetName  表格索引名 
     * @return 
     */  
    public static boolean sheetExist(String fileDir,String sheetName) throws Exception{  
         boolean flag = false;  
         File file = new File(fileDir);  
         if(file.exists()){    //文件存在  
            //创建workbook  
             try {  
                workbook = new HSSFWorkbook(new FileInputStream(file));  
                //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)  
                HSSFSheet sheet = workbook.getSheet(sheetName);    
                if(sheet!=null)  
                    flag = true;  
            } catch (Exception e) {  
                throw e;
            }   
              
         }else{    //文件不存在  
             flag = false;  
         }  
         return flag;  
    }  
    /** 
     * 创建新excel. 
     * @param fileDir  excel的路径 
     * @param sheetName 要创建的表格索引 
     * @param titleRow excel的第一行即表格头 
     */  
    public static void createExcel(String fileDir,String sheetName,String titleRow[]){  
        //创建workbook  
        workbook = new HSSFWorkbook();  
        //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)  
        HSSFSheet sheet1 = workbook.createSheet(sheetName);    
        //新建文件  
        FileOutputStream out = null;  
        try {  
            //添加表头   
            HSSFRow row = workbook.getSheet(sheetName).createRow(0);    //创建第一行    
            for(short i = 0;i < titleRow.length;i++){  
                HSSFCell cell = row.createCell(i);  
                cell.setCellValue(titleRow[i]);  
            }  
            out = new FileOutputStream(fileDir);  
            workbook.write(out);  
        } catch (Exception e) {  
        	 e.printStackTrace();
        } finally {    
            try {    
                out.close();    
            } catch (IOException e) {    
                e.printStackTrace();  
            }    
        }    
    }  
    /** 
     * 删除文件. 
     * @param fileDir  文件路径 
     */  
    public static boolean deleteExcel(String fileDir) {  
        boolean flag = false;  
        File file = new File(fileDir);  
        // 判断目录或文件是否存在    
        if (!file.exists()) {  // 不存在返回 false    
            return flag;    
        } else {    
            // 判断是否为文件    
            if (file.isFile()) {  // 为文件时调用删除文件方法    
                file.delete();  
                flag = true;  
            }   
        }  
        return flag;  
    }  
    /** 
     * 往excel中写入(已存在的数据无法写入). 
     * @param fileDir    文件路径 
     * @param sheetName  表格索引 
     * @param object 
     * @throws Exception 
     */  
    public static void writeToExcel(String fileDir,String sheetName,List<Map> mapList){  
        //创建workbook  
        File file = new File(fileDir);  
        try {  
            workbook = new HSSFWorkbook(new FileInputStream(file));  
        } catch (FileNotFoundException e) {  
            e.printStackTrace();  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
        //流  
        FileOutputStream out = null;  
        HSSFSheet sheet = workbook.getSheet(sheetName);  
        // 获取表格的总行数  
         int rowCount = sheet.getLastRowNum() + 1; // 需要加一  
        // 获取表头的列数  
        int columnCount = sheet.getRow(0).getLastCellNum();  
        try {   
            // 获得表头行对象  
            HSSFRow titleRow = sheet.getRow(0);  
            if(titleRow!=null){ 
                for(int rowId=0;rowId<mapList.size();rowId++){
                    Map map = mapList.get(rowId);
                    HSSFRow newRow=sheet.createRow(rowId+rowCount);
                    for (short columnIndex = 0; columnIndex < columnCount; columnIndex++) {  //遍历表头  
                        String mapKey = titleRow.getCell(columnIndex).toString().trim().toString().trim();  
                        HSSFCell cell = newRow.createCell(columnIndex);  
                        cell.setCellValue(map.get(mapKey)==null ? null : map.get(mapKey).toString());  
                    } 
                }
            }  
  
            out = new FileOutputStream(fileDir);  
            workbook.write(out);  
        } catch (Exception e) {  
             e.printStackTrace();
        } finally {    
            try {    
                out.close();    
            } catch (IOException e) {    
                e.printStackTrace();  
            }    
        }    
    }  
      
    public static void main(String[] args) {  
    	//读文件
    	String pathRead = "C:/Users/car/Desktop/gaode/order_all.xlsx";
    	String titleRead[] = {"order_no","order_id","actual_board_time","vehicle_no","actual_off_time"};
    	String pathWrite = "C:/Users/car/Desktop/gaode/2.xls";
    	String titleWrite[] = {"order_id","order_no","vehicle_no","polestar"};
    	try {
    		ArrayList<HashMap<String, String>> result = readXlsx(pathRead, titleRead);
    		System.out.println(result.size());
    		System.out.println(result.get(0).get("order_no"));
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
    	
        //写文件
    	List<Map> list=new ArrayList<Map>();
        Map<String,String> map=new HashMap<String,String>();
        map.put(titleWrite[0], "111");
        map.put(titleWrite[1], "张三");
        map.put(titleWrite[2], "111！@#");
        map.put(titleWrite[3], "polestar@");
        
        list.add(map);
        list.add(map);
        writeXls(list, titleWrite, pathWrite);
    }  
    
    public static void writeXls(List<Map> list,String title[],String path){
    	System.out.println("writeXls");
         if(!fileExist(path)){
     		try {
 				createExcel(path,"sheet1",title);
 			} catch (Exception e) {
 				// TODO Auto-generated catch block
 				e.printStackTrace();
 			}
     	}
         try {
 			writeToExcel(path,"sheet1",list);
 		} catch (Exception e) {
 			// TODO Auto-generated catch block
 			e.printStackTrace();
 		} 
    }
    
    static SimpleDateFormat sdf1=new SimpleDateFormat("yyyymmddHHmmss");
    public static ArrayList<HashMap<String, String>> readXlsx(String path,String[] readTitle) throws FileNotFoundException {
    	System.out.println("readXlsx");
    	ArrayList<HashMap<String, String>> list = new ArrayList<>();
    	try {
    		org.apache.poi.xssf.usermodel.XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(path));// 得到这个excel表格对象
    		XSSFSheet sheet = work.getSheetAt(0); //得到第一个sheet
    		int rowNo = sheet.getLastRowNum(); //得到行数
    		for (int i = 1; i < rowNo; i++) {
    			
	    		XSSFRow row = sheet.getRow(i);
	    		HashMap<String, String> map = new HashMap<>();
	    		for(int j = 0;j < readTitle.length;j++){//得到列数字
	    			XSSFCell cell = row.getCell(j);
	    			if(cell!=null){
	    				String ce1 = "";
	    				if(Cell.CELL_TYPE_NUMERIC == cell.getCellType()){
	    					
	    					 if(HSSFDateUtil.isCellDateFormatted(cell)){
	    					       Date d = (Date) cell.getDateCellValue();
	    					       ce1 =  sdf1.format(d);
	    					 }else{
	    					       //使用DecimalFormat对double进行了格式化，随后使用format方法获得的String就是你想要的值了。
	    					      DecimalFormat df = new DecimalFormat("0");
	    					      ce1 = String.valueOf(df.format(cell.getNumericCellValue()));
	    					}
	    				}else{
	    					cell.setCellType(Cell.CELL_TYPE_STRING);
	    					ce1 = cell.getStringCellValue();
	    				}
	    			
//	    				System.out.println(ce1 );
	    				map.put(readTitle[j], ce1);
	    			}
	    		}
	    		
	    		list.add(map);
    		}
		} catch (Exception e) {
			e.printStackTrace();
		}
    	return list;
    }

    
    
}