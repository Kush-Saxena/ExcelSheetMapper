package constants;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class InitParameters {

	public static Map<String, String> colNamesMap = null;

	public static List<String> scenarioNumbers = null;

	public static String inpFilePath = null;

	public static String outFilePath = null;

	public static String sheetName = null;

	public static int scenarioRow;

	public static int scenarioCol;

	static {
		colNamesMap = new HashMap<String, String>();
		colNamesMap.put("ScenarioNo.", "SCNo.");
		colNamesMap.put("ColA", "Col1");
		colNamesMap.put("ColC", "Col2");

		scenarioNumbers = new ArrayList<String>();
		scenarioNumbers.add("1.0");
		scenarioNumbers.add("3.0");
		scenarioNumbers.add("5.0");

		inpFilePath = "C:/Users/kusaxena/OneDrive - Capgemini/Documents/Projects/Book1.xls";
		outFilePath = "C:/Users/kusaxena/OneDrive - Capgemini/Documents/Projects/Book2.xls";

		sheetName = "Sheet1";

		scenarioRow = 0;
		scenarioCol = 0;
	}
}
