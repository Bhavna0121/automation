package com.sample.mongodb;

import org.bson.Document;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;

public class DropingCollection {
	public static void main(String[] args) {
		MongoClient mongo = new MongoClient("localhost", 27017);

		// Creating Credentials
		MongoCredential credential;
		credential = MongoCredential.createCredential("sampleUser", "myDb", "password".toCharArray());
		System.out.println("Connected to the database successfully");

		// Accessing the database
		MongoDatabase database = mongo.getDatabase("myDb");

		// Creating a collection
		System.out.println("Collections created successfully");

		// Retieving a collection
		MongoCollection<Document> collection = database.getCollection("sampleCollection");

		// Dropping a Collection
		collection.drop();
		System.out.println("Collection dropped successfully");
	}

}
