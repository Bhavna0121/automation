package com.sample.Classes;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

public class ThreadExample {
	Log logger = LogFactory.getLog(getClass());
	
	// public static void main(String[] args) {
	 String st1 = "ABC";
	 String st2 = "XYZ";

	Thread t = new Thread("Thread 1") {

		public void run() {
			while (true) {
				synchronized (st1) {
					synchronized (st2) {
						logger.info(st1 + " " + st2);
					}

				}
			}
		}

		// }

	};

	Thread t2 = new Thread("Thread 2") {

		public void run() {
			while (true) {
				synchronized (st2) {
					synchronized (st1) {
						System.out.println(st2 + " " + st1);
					}

				}
			}
		}

		// }

	};

	public static void main(String[] args) {
		ThreadExample threadEx = new ThreadExample();
		threadEx.t.start();
		threadEx.t2.start();
	}
}
