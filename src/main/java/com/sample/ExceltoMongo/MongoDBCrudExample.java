package com.sample.ExceltoMongo;


import java.util.concurrent.ConcurrentHashMap;
import com.mongodb.BasicDBObjectBuilder;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;
import com.mongodb.WriteResult;

public class MongoDBCrudExample {
	public static ConcurrentHashMap<String, Object> workbookMap = new ConcurrentHashMap<String, Object>();
	public void crud(String path, int count) throws Exception {
		String filename = path.substring((path.lastIndexOf("/")+1), path.length());
		User user = createUser(filename,count);
		DBObject doc = createDBObject(user);
		
		MongoClient mongo = new MongoClient("localhost", 27017);
		DB db = mongo.getDB("ScenarioFile");
		
		DBCollection col = db.getCollection("Scenariotable");
	//	db.collection.findOne({_id: "myId"}, {_id: 1});
		//create user
		WriteResult result = col.insert(doc);
//		System.out.println(result.getUpsertedId());
//		System.out.println(result.getN());
//		System.out.println(result.isUpdateOfExisting());
		//System.out.println(((Object) result).getLastConcern());
		
		//read example
		DBObject query = BasicDBObjectBuilder.start().add("_id", user.getId()).get();
		DBCursor cursor = col.find(query);
		while(cursor.hasNext()){
			System.out.println(cursor.next());
		}
		
		Xls_Reader xls = new Xls_Reader(path,workbookMap);
		System.out.println(xls);// "AdvanceCompletion.xls"
		ConcurrentHashMap<String, Object> testDataSheet = (ConcurrentHashMap<String, Object>) workbookMap.get(filename);

		String stringtestDataSheet = testDataSheet.toString();
		//update example
		user.setScenarioData(stringtestDataSheet);
		doc = createDBObject(user);
		result = col.update(query, doc);
		DBObject query1 = BasicDBObjectBuilder.start().add("_id", user.getId()).get();
		DBCursor cursor1 = col.find(query1);
		while(cursor1.hasNext()){
			System.out.println(cursor1.next());
		}
//		System.out.println(result.getUpsertedId());
//		System.out.println(result.getN());
//		System.out.println(result.isUpdateOfExisting());
		//System.out.println(result.getLastConcern());
		
//		//delete example
//		result = col.remove(query);
//		System.out.println(result.getUpsertedId());
//		System.out.println(result.getN());
//		System.out.println(result.isUpdateOfExisting());
		//System.out.println(result.getLastConcern());
		
		//close resources
		mongo.close();
	}

	private static DBObject createDBObject(User user) {
		BasicDBObjectBuilder docBuilder = BasicDBObjectBuilder.start();
								
		docBuilder.append("_id", user.getId());
		docBuilder.append("name", user.getScenarioName());
		docBuilder.append("Scenario Data", user.getScenarioData());
//		docBuilder.append("role", user.getRole());
//		docBuilder.append("isEmployee", user.isEmployee());
		return docBuilder.get();
	}

	private static User createUser(String path, int count) {
		return null;
//		User u = new User();
//		u.setId(count);
//		u.setScenarioName(path);
//		String testDataSheet = null;
//		//ConcurrentHashMap<String, Object> testDataSheet = null;
//		u.setScenarioData(testDataSheet);
//		return u;
	}

}
