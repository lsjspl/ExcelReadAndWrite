package com.master5.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class Start {

	static Logger logger = Logger.getLogger(Start.class);

	// 第一行为标题

	static List<String> paths;

	static String path;

	static List<Map<String, Object>> datas;

	static ArrayList<Map<String, Object>> results;
	static String[] titles;

	static String[] resultTitles = { "姓名", "考勤号码", "日期", "首次打卡", "末次打卡", "结果", "备注" };
	static HSSFWorkbook hssfWorkbook;

	// 注意安全
	static SimpleDateFormat yearFormat = new SimpleDateFormat("yyyy-MM-dd");
	static SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm");
	static SimpleDateFormat monthFormat = new SimpleDateFormat("yyyy-MM");
	static SimpleDateFormat dayFormat = new SimpleDateFormat("dd");

	public static void main(String[] args) {

		logo();

		logger.info("添加了新任务：" + Calendar.getInstance().getTime());

		init();

		if (paths.isEmpty()) {
			logger.info("没有文件啊");
			return;
		}
		for (String pathTmp : paths) {
			path = pathTmp;
			System.out.println("--------------------------------------------------------------------------------------");
			logger.info("处理：" + path);
			try {

				saxExcel();

				HandlerData();

				writeExcel();

				logger.info("处理成功：" + path);
				logger.info("输出结果：" + path.replaceAll(".xls|.XLS", "_result.xls"));
			} catch (Exception e) {
				e.printStackTrace();
				logger.error("处理失败 excel格式不正确：" + path);
			}

		}

		logger.info("任务处理完成：" + Calendar.getInstance().getTime());
	}

	private static void logo() {

		System.out.println("-----------------------------------------------------");
		System.out.println("--             考勤数据处理应用                    --");
		System.out.println("--                                                 --");
		System.out.println("--                        Ver:1.02 Alpha           --");
		System.out.println("--                        Powered  By 五少爷       --");
		System.out.println("-----------------------------------------------------");
	}

	private static void init() {

		paths = new ArrayList<>();

		File file = new File("./");

		logger.info("正在扫描如下目录：" + file.getAbsolutePath());

		String[] fileNames = file.list();

		for (String name : fileNames) {
			if (!name.contains("_result") && name.toLowerCase().endsWith(".xls")) {
				paths.add(name);
				logger.info("已提取excel文件：" + name);
			}
		}
	}

	@SuppressWarnings("deprecation")
	private static void saxExcel() throws FileNotFoundException, IOException {

		datas = new ArrayList<>();

		hssfWorkbook = new HSSFWorkbook(new FileInputStream(path));

		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}

			int rowLength = hssfSheet.getLastRowNum();
			int cellLength = 0;

			HSSFRow titleRow = hssfSheet.getRow(0);
			if (titleRow == null) {
				continue;
			}
			cellLength = titleRow.getLastCellNum();
			titles = new String[cellLength];
			for (int i = 0; i < cellLength; i++) {
				titles[i] = titleRow.getCell(i).getStringCellValue();
			}

			for (int rowNum = 1; rowNum < rowLength; rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow != null) {
					Map<String, Object> map = new HashMap<>();
					for (int i = 0; i < cellLength; i++) {
						Cell cell = hssfRow.getCell(i);
						if (cell == null) {
							map.put(titles[i], null);
							continue;
						}

						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							map.put(titles[i], cell.getStringCellValue());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							if (HSSFDateUtil.isCellDateFormatted(cell)) {
								map.put(titles[i], cell.getDateCellValue());
							} else {
								map.put(titles[i], cell.getNumericCellValue());
							}
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							map.put(titles[i], cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							map.put(titles[i], cell.getCellFormula());
							break;
						case Cell.CELL_TYPE_BLANK:
							map.put(titles[i], "");
							break;
						default:
							break;
						}

					}

					datas.add(map);
				}

			}

		}

	}

	private static void HandlerData() throws ParseException {
		Date laterLimitFirst = timeFormat.parse("09:00");
		Date laterLimit = timeFormat.parse("09:30");
		String name;
		String number;
		Date date = null;
		String dateStr;
		Calendar calendar;
		String timeStr;
		long hour = 0;
		long minutes = 0;

		String token;

		TreeSet<String> numberSet = new TreeSet<>();
		Map<String, Object> resultElement;
		Map<String, String> number2Name = new HashMap<>();

		// 整理数据
		Map<String, Map<String, Object>> total = new HashMap<>();
		for (Map<String, Object> map : datas) {

			name = (String) map.get("姓名");
			if (name == null) {
				continue;
			}
			number = (String) map.get("考勤号码");
			date = (Date) map.get("日期时间");
			dateStr = yearFormat.format(date);
			timeStr = timeFormat.format(date);

			if (!number2Name.containsKey(number)) {
				number2Name.put(number, name);
			}

			token = new StringBuilder().append(number).append(",").append(Integer.valueOf(dayFormat.format(date))).toString();

			if (!total.containsKey(token)) {
				resultElement = new HashMap<>();
				resultElement.put("姓名", name);
				resultElement.put("考勤号码", number);
				resultElement.put("日期", dateStr);
				resultElement.put("首次打卡", timeStr);
				resultElement.put("末次打卡", "");
				resultElement.put("结果", "");
				resultElement.put("备注", "");
				total.put(token, resultElement);
			} else {

				resultElement = total.get(token);
				Date now = timeFormat.parse(timeStr);
				Date upTmp = timeFormat.parse((String) resultElement.get("首次打卡"));
				Date downTmp = ((String) resultElement.get("末次打卡")).equals("") ? null : timeFormat.parse((String) resultElement.get("末次打卡"));
				if (upTmp.after(now)) {
					resultElement.put("首次打卡", timeStr);
				} else if (downTmp == null || downTmp.before(now)) {
					resultElement.put("末次打卡", timeStr);
				}

			}

			// 处理迟到信息

			date = yearFormat.parse((String) resultElement.get("日期"));
			Date firstSignIn = timeFormat.parse((String) resultElement.get("首次打卡"));
			Date lastSignIn = ((String) resultElement.get("末次打卡")).equals("") ? null : timeFormat.parse((String) resultElement.get("末次打卡"));
			String remarks = "";
			long later = 0;
			long leave = 0;
			if (lastSignIn != null) {
				later = (firstSignIn.getTime() - laterLimit.getTime()) / (1000 * 60);
				Date checkMoring = firstSignIn.before(laterLimitFirst) ? laterLimitFirst : firstSignIn.after(laterLimit) ? laterLimit : firstSignIn;

				leave = (9 * 1000 * 60 * 60 - (lastSignIn.getTime() - checkMoring.getTime())) / (1000 * 60);

			}

			if (later > 0) {
				hour = later / 60;
				minutes = later % 60;
				remarks += "迟到" + (hour > 0 ? hour + "小时" : "") + (minutes > 0 ? minutes + "分钟" : "") + ";";

				resultElement.put("结果", remarks);
			} else {
				resultElement.put("结果", "");
			}

			remarks = "";
			if (leave > 0) {
				hour = leave / 60;
				minutes = leave % 60;
				remarks += "早退" + (hour > 0 ? hour + "小时" : "") + (minutes > 0 ? minutes + "分钟" : "") + ";";

				resultElement.put("备注", remarks);
			} else {
				resultElement.put("备注",lastSignIn==null?"少打一次卡": "");
			}

			numberSet.add(number);
		}

		calendar = Calendar.getInstance();
		calendar.setTime(date);
		int maxDays = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
		int month = calendar.get(Calendar.MONTH);
		int year = calendar.get(Calendar.YEAR);
		boolean isWeek;
		for (int day = 1; day <= maxDays; day++) {
			for (String numberTmp : numberSet) {
				token = numberTmp + "," + day;
				calendar.set(year, month, day);
				isWeek=calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY || calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY;
				if (!total.containsKey(token) && !isWeek) {
					resultElement = new HashMap<>();
					resultElement.put("姓名", number2Name.get(numberTmp));
					resultElement.put("考勤号码", numberTmp);
					resultElement.put("日期", yearFormat.format(calendar.getTime()));
					resultElement.put("首次打卡", "");
					resultElement.put("末次打卡", "");
					resultElement.put("结果", "");
					resultElement.put("备注", "今天没打卡");
					total.put(token, resultElement);
				}
			}
		}

		results = new ArrayList<>(total.values());
		results.sort((Map<String, Object> map1, Map<String, Object> map2) -> {

			int no1 = Integer.valueOf((String) map1.get("考勤号码"));
			int no2 = Integer.valueOf((String) map2.get("考勤号码"));
			String day1 = (String) map1.get("日期");
			String day2 = (String) map2.get("日期");

			if (no1 == no2) {
				try {
					return yearFormat.parse(day1).compareTo(yearFormat.parse(day2));
				} catch (ParseException e) {
					e.printStackTrace();
				}
			}

			return no1 - no2;
		});

	}

	private static void writeExcel() throws IOException {

		HSSFSheet sheet = hssfWorkbook.createSheet("汇总");

		HSSFRow title = sheet.createRow(0);

		for (int cellIndex = 0; cellIndex < resultTitles.length; cellIndex++) {
			title.createCell(cellIndex).setCellValue(resultTitles[cellIndex]);
		}

		for (int rowIndex = 0; rowIndex < results.size(); rowIndex++) {
			Map<String, Object> map = results.get(rowIndex);
			HSSFRow row = sheet.createRow(rowIndex + 1);
			for (int cellIndex = 0; cellIndex < resultTitles.length; cellIndex++) {
				row.createCell(cellIndex).setCellValue(map.get(resultTitles[cellIndex]).toString());
			}
		}

		hssfWorkbook.write(new File(path.replaceAll(".xls|.XLS", "_result.xls")));

	}

}
