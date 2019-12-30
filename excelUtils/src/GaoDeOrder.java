

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class GaoDeOrder {
	public static void main(String[] args) {
      //创建文件  
      String readTitle[] = {"order_id","order_no","actual_board_time","vehicle_no","actual_off_time"};
      String writeTitle[] = {"order_id","order_no","abord_time","off_time","polestarPoint","polestarMile"};
      String path = "C:/Users/car/Desktop/gaode/order_all.xlsx";
      
      if(!ExcelUtils.fileExist(path)){
    	  System.out.println(path +"不存在");
    	  return;
      }
      
      ArrayList<HashMap<String, String>> data = new  ArrayList<HashMap<String, String>>();
      try {
    	  
    	  data = ExcelUtils.readXlsx(path,readTitle);
      	 } catch (Exception e1) {
      		 System.out.println(e1.getMessage());
		e1.printStackTrace();
      }
      
      for(int i = 0;i < data.size();i++){
    	  HashMap<String, String> item = data.get(i);
    	  System.out.println(item.get(readTitle[0])+"  "+item.get(readTitle[1])
    			  +"  "+item.get(readTitle[2])+"  "+item.get(readTitle[3])
    			  +"  "+item.get(readTitle[4]));
		}
	}

	
}
