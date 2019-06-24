package com.sample.ExceltoMongo;

import java.util.concurrent.ConcurrentHashMap;

public class User implements Comparable<User> {

	private int id;
	private String ScenarioName;
	private String ScenarioData;

	public String getScenarioData() {
		return ScenarioData;
	}

	public void setScenarioData(String testDataSheet) {
		ScenarioData = testDataSheet;
	}

	// private String role;
	// private boolean isEmployee;
	public int getId() {
		return id;
	}
	
	public User(String name, int id ) {
		this.id = id;
		this.ScenarioName = name;
		// TODO Auto-generated constructor stub
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getScenarioName() {
		return ScenarioName;
	}

	public void setScenarioName(String scenarioName) {
		ScenarioName = scenarioName;
	}

	public int compareTo(User o) {
		return id - o.id;
	}

}
