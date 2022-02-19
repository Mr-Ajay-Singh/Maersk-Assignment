import java.util.Comparator;

import org.graalvm.compiler.lir.StandardOp.ImplicitNullCheck;

public class ExcelDataStorage implements Comparator<ExcelDataStorage> {
	private int id;
	private String name;
	private long salary;
	public ExcelDataStorage() {
	}
	
	ExcelDataStorage(int id, String name, long salary){
		this.id = id;
		this.name = name;
		this.salary = salary;
	}
	
	public int compare(ExcelDataStorage excelDataStorage1, ExcelDataStorage excelDataStorage2) {
		if(excelDataStorage1.salary< excelDataStorage2.salary) {
			return 1;
		}else if(excelDataStorage1.salary == excelDataStorage2.salary) {
			return 0;
		}else {
			return -1;
		}
	}
	int getId() {
		return id;
	}
	String getName() {
		return name;
	}
	long getSalary() {
		return salary;
	}
	
}
