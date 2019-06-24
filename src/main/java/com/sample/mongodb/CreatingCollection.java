package com.sample.mongodb;


import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.client.MongoDatabase;

public class CreatingCollection {
	public static void main(String[] args) {
		
	
	MongoClient mongo = new MongoClient( "localhost" , 27017 ); 
	 // Creating Credentials 
    MongoCredential  credential = MongoCredential.createCredential("sampleUser", "myDb", 
       "password".toCharArray()); 
    System.out.println("Connected to the database successfully"); 
    //Accessing the database 
    MongoDatabase database = mongo.getDatabase("myDb");  
    //Creating a collection 
    database.createCollection("sampleCollection1"); 
    System.out.println("Collection created successfully");
}
}

