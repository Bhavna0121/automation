package com.sample.Scenario;

import java.io.File;
import java.util.LinkedHashMap;
import java.util.Map;

import com.sample.ExceltoMongo.MongoDBCrudExample;
import com.sample.ExceltoMongo.Xls_Reader;
import com.sample.mongodb.InsertingDocument;

public class ScenarioTable {
	public static void main(String[] args) {
		File folder = new File("/Users/bhavnajain/Downloads/Scenarios");
		File[] listOfFiles = folder.listFiles();
		int count = 0;
		for (File file : listOfFiles) {
			if (file.getName().startsWith(".")) {
				continue;
			}
			try {
				if (file.isFile()) {
					count ++;
					if (file.getName().endsWith(".xls")) {
					//	Xls_Reader xls = new Xls_Reader(folder + "/" + file.getName());
//						InsertingDocument i = new InsertingDocument();
//						i.insert(folder + "/" + file.getName(),count);
						MongoDBCrudExample m = new MongoDBCrudExample();
						m.crud(folder + "/" + file.getName(),count);
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			}

		}
	}
}
