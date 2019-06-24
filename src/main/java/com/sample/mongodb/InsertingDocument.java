package com.sample.mongodb;

import java.rmi.UnknownHostException;

import org.bson.Document;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.sample.ExceltoMongo.Xls_Reader;
import com.sample.Scenario.KeyExistsInMongo;

public class InsertingDocument {
	public void insert(String path, int count) throws UnknownHostException {
		MongoClient mongo = new MongoClient("localhost", 27017);

		// Creating Credentials
		MongoCredential credential;
		credential = MongoCredential.createCredential("sampleUser", "myDb", "password".toCharArray());
		System.out.println("Connected to the database successfully");

		// Accessing the database
		MongoDatabase database = mongo.getDatabase("myDb");
		path = path.substring((path.lastIndexOf("/")+1), path.length());
		// Retrieving a collection
		MongoCollection<Document> collection = database.getCollection("sampleCollection");
		System.out.println("Collection sampleCollection selected successfully");
		Document document = new Document("ScenarioName", path).append("id", count).append("description", "database");
		collection.insertOne(document);
		System.out.println("Document inserted successfully");
		MongoCollection<Document> collection1 = database.getCollection("sampleCollection1");
		System.out.println("Collection sampleCollection1 selected successfully");
		KeyExistsInMongo.existstrue(path);
	}

}
