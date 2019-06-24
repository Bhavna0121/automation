package Sample.Sample;

import java.util.Comparator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import com.sample.ExceltoMongo.User;

public class TreeMapSample {
	
	public static void main(String[] args) {
		Map<User, String> map = new TreeMap<User, String>();
		map.put(new User("Ram",3000), "RAM");
		map.put(new User("John",6000), "JOHN");
		map.put(new User("Crish",2000), "CRISH");
		map.put(new User("Tom",2400), "TOM");
		Set<User> keys = map.keySet();
        for(User key:keys){
            System.out.println(key+" ==> "+map.get(key));
        }
        System.out.println("===================================");
	}
}
