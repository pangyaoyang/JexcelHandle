package com.pangyy.jexcelHandle;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.pangyy.jexcelHandle.common.ExcelUtil;
import com.pangyy.jexcelHandle.common.WriteFile;


public class CreateSqlMain 
{
	static public String  outFileDir = "E:/alipay/上线sql/codeout/"; 
	
    public static void main( String[] args )
    {
    	// insert app
//    	insertApp();
    }
    
    
    static public void insertApp() {
    	String outSqlFileName = outFileDir + "insert_app_v1.sql";
        List<Object>   cityResult = ExcelUtil.readXlsxFileToArray("E:/alipay/城市服务基础应用信息.xls");
        
        if(cityResult == null){
        	System.out.println("read  excel failed");
        	return;
        }
        
        System.out.println("excel date row conut:" + cityResult.size());

        String  appSql = "INSERT INTO industry_app  (app_id ,app_name ,app_status,app_logo,app_type )  values  \n";
        for(Object ob:cityResult) {
        	/**
        	 * 取一行数据
        	 */
            @SuppressWarnings("unchecked")
			ArrayList<String> arr = (ArrayList<String>) ob;

            /**
             * 取特定列的值
             */
            String app_id = arr.get(1);
            String app_name = arr.get(3);
            String app_type = arr.get(4);

            if(StringUtils.equals(app_type,"SREVICE_WINDOW")){
                app_type = "windowserver";
            }else{
                app_type = "aliph5";
            }
            String app_logo = arr.get(7);

            appSql += " (\"" + app_id + "\", ";
            appSql += "\"" + app_name + "\", ";
            appSql += "\"" + "ONLINE" + "\", ";
            appSql += "\"" + app_type + "\", ";
            appSql += "\"" + app_logo +  "\"),\n";
        }

        System.out.println(appSql);
        
        WriteFile.writeFile(outSqlFileName, appSql);
    }

}
