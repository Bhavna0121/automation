package com.sample.Scenario;


import java.rmi.UnknownHostException;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;

public class KeyExistsInMongo {
	public static void existstrue(String path) throws UnknownHostException {
		 
        // Get a db connection
        MongoClient m1 = new MongoClient("localhost");
 
        // connect to test db,use your own here
        DB db = m1.getDB("test");
 
        // obtain the car collection
        DBCollection coll = db.getCollection("sampleCollection");
 
        BasicDBObject query = new BasicDBObject();
        query.put("ScenarioName", new BasicDBObject("ScenarioName", path));
 
        // store the documents in cursor car
        DBCursor car = coll.find(query);
 
        // iterate and print the contents of cursor
        try {
            while (car.hasNext()) {
                System.out.println(car.next());
            }
        } finally {
            // close the cursor
            car.close();
        }
 
    }

}
