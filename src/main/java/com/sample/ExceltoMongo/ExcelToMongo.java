package com.sample.ExceltoMongo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

import com.mongodb.BasicDBObject;
import com.mongodb.DBCollection;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;

public class ExcelToMongo {
	public static void main(String[] args) {
		MongoClient mongo = new MongoClient( "localhost" , 27017 );
		 MongoDatabase database = mongo.getDatabase("myDb"); 
        System.out.println("Connected to Database successfully");
        MongoCollection<Document> collection = database.getCollection("sampleCollectionExcelData");
        System.out.println("Collection your_collection name selected successfully");

//             DBCollection OR = database.getCollection("Input_Container");
//             System.out.println("Collection Device_Details selected successfully");
             collection.drop();
             DBObject arg1 = null;
             //coll.update(query, update);
             database.createCollection("sampleCollectionExcelData"); 
             String path ="/Users/bhavnajain/Desktop/Workbook1.xlsx";

             File myFile = new File(path);
             FileInputStream inputStream = null;
			try {
				inputStream = new FileInputStream(myFile);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
             XSSFWorkbook workbook = null;
			try {
				workbook = new XSSFWorkbook(inputStream);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
             int number=workbook.getNumberOfSheets();
               System.out.println("NumberOfSheets "+number);

               for(int i=0;i<number;i++)
               {
             XSSFSheet sheet = workbook.getSheetAt(i);
             int col_value=sheet.getRow(0).getLastCellNum();
             int row_num= sheet.getLastRowNum();
             System.out.println("row_num "+row_num);
             List<String> DBheader = new ArrayList<String>();
             List<String> Data = new ArrayList<String>();

             for(int z=1;z<=row_num;z++){
                  DBheader.clear();
                  Data.clear();
              for(int j=0;j<col_value;j++)
             {
                 if(sheet.getRow(0).getCell(j).toString()!=null || sheet.getRow(0)!=null)
                 {
                 String cel_value = sheet.getRow(0).getCell(j).toString();
                 DBheader.add(cel_value.trim());
                 }
                 else{
                     break;

                 }
             }
             for(int k=0;k<col_value;k++){
                 String data =" ";   
                 if(sheet.getRow(z).getCell(k)!=null)
                 {
                 data =  sheet.getRow(z).getCell(k).toString();
                 }
                 Data.add(data.trim());

                 }
             BasicDBObject doc = new BasicDBObject();
             System.out.println("Data.size() "+Data.size());

             int l=0;
             for(String headers:DBheader)
             { 
             if(l>Data.size()){break;}
                 doc.append(headers, Data.get(l));
                 l++;
              }
             ((DBCollection) database).insert(doc);
             }

         }System.out.println("File Upload Done");
         mongo.close();
	}

}
