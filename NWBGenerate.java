import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NWBGenerateMain {

	public static void main(String[] args) {
		// Parameters Setting
		String dirPath = "C:\\Users\\wangh\\OneDrive\\WOS\\CS2014-2019";
		boolean containTitle = true;
		int indexOfKeywords = 4; // from 0
		String seperator = ";";

		nwbGenerateByExcel(dirPath, containTitle, indexOfKeywords, seperator);
	}

	public static void nwbGenerateByExcel(String dirPath, boolean containTitle, int indexOfKeywords, String seperator) {
		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		XSSFRow row = null;
		PrintStream ps = null;

		String[] keywords = null;
		Map<String, Integer> map = null;
		Map<String, Integer> comap = null;
		List<String[]> list = null;

		File dir = new File(dirPath);
		for (File f : dir.listFiles()) {
			map = new TreeMap<>();
			comap = new TreeMap<>();
			list = new ArrayList<>();
			System.out.println("Processing: " + f.getName());
			try {
				wb = new XSSFWorkbook(new FileInputStream(f));
			} catch (IOException e) {
				e.printStackTrace();
			}
			sheet = wb.getSheetAt(0);
			for (int i = containTitle ? 1 : 0; i <= sheet.getLastRowNum(); i++) {
				row = sheet.getRow(i);
				if(row.getCell(indexOfKeywords).getStringCellValue().trim().isEmpty())
					continue;
				keywords = row.getCell(indexOfKeywords).getStringCellValue().trim().toLowerCase().split(seperator);
				List<String> keywords_list = new ArrayList<String>();
				for (String s : keywords) {
					if (!(s.trim().isEmpty()))
						keywords_list.add(s.trim());
				}

				for (int j = 0; j < keywords_list.size(); j++) {
					if (!map.containsKey(keywords_list.get(j))) {
						map.put(keywords_list.get(j), list.size());
						list.add(new String[] { keywords_list.get(j), "1" });
					} else {
						list.get(map.get(keywords_list.get(j)))[1] = (Integer
								.parseInt(list.get(map.get(keywords_list.get(j)))[1]) + 1 + "");
					}
					for (int k = j + 1; k < keywords_list.size(); k++) {
						// check whether keywords[k] is contained by map
						if (!map.containsKey(keywords_list.get(k))) {
							map.put(keywords_list.get(k), list.size());
							list.add(new String[] { keywords_list.get(k), "0" });
						}

						String key = null;
						if (map.get(keywords_list.get(j)) < map.get(keywords_list.get(k))) {
							key = (100001 + map.get(keywords_list.get(j))) + " "
									+ (100001 + map.get(keywords_list.get(k)));
						} else {
							key = (100001 + map.get(keywords_list.get(k))) + " "
									+ (100001 + map.get(keywords_list.get(j)));
						}
						if (!comap.containsKey(key)) {
							comap.put(key, 1);
						} else {
							comap.put(key, comap.get(key) + 1);
						}
					}
				}
			}

			try {
				File outputFile = new File(dir.getParent() + "/results");
				if (!outputFile.exists()) {
					outputFile.mkdir();
				}
				ps = new PrintStream(dir.getParent() + "/results/" + f.getName().split(".")[0] + ".nwb");
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}

			StringBuilder sb = new StringBuilder("");
			// remove the nodes only appear once or in stopwords
			List<String> removeList = new ArrayList<>();
			for (int i = 0; i < list.size(); i++) {
				if ((list.get(i)[1]).equals("1") || isStopword(list.get(i)[0])) {
					removeList.add(100001 + i + "");
				} else {
					sb.append(100001 + i + " \"").append(list.get(i)[0]).append("\" ").append(list.get(i)[1])
							.append(System.getProperty("line.separator"));
				}
			}

			ps.println("*Nodes " + (list.size() - removeList.size()));
			ps.println("id*int lable*string weight*int");
			ps.print(sb.toString());

			sb = null;
			sb = new StringBuilder("");
			int edgeNum = 0;
			for (Entry<String, Integer> s : comap.entrySet()) {
				boolean flag = true;
				for (int i = 0; i < removeList.size(); i++) {
					if (s.getKey().contains(removeList.get(i))) {
						flag = false;
						break;
					}
				}
				if (flag) {
					sb.append(s.getKey()).append(" ").append(s.getValue()).append(System.getProperty("line.separator"));
					edgeNum++;
				}
			}

			ps.println("*UndirectedEdges " + edgeNum);
			ps.println("source*int target*int weight*int");
			ps.print(sb.toString());
			sb = null;
			map = null;
			comap = null;
			list = null;
			ps.close();
			System.out.println("Complete: " + f.getName());
			System.gc();
		}

	}

	public static boolean isStopword(String word) {
		boolean flag = false;
		String[] stopwords = { "" };
		for (String s : stopwords) {
			if (s.equals(word)) {
				flag = true;
				break;
			}
		}
		return flag;
	}

}
