package edu.abhi.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author abhishek sarkar
 *
 */
public class ExcelDiffChecker {

	private static boolean diffFound = false;
	private static boolean commentFlag = true;
	private static String FILE_NAME1, FILE_NAME2;

	public static void main(String[] args) {

		FILE_NAME1 = args[0];
		FILE_NAME2 = args[1];
		commentFlag = args.length == 2;

		String RESULT_FILE = FILE_NAME1.substring(0, FILE_NAME1.lastIndexOf(".")) + " vs " + FILE_NAME2;

		File resultFile = new File(RESULT_FILE);

		Utility.deleteIfExists(resultFile);

		try(XSSFWorkbook resultWorkbook = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME1)));
				XSSFWorkbook workbook2 = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME2)))) {

			diffFound = CallableValue.analyzeResult(processAllSheets(resultWorkbook, workbook2));

			if(diffFound) {
				if(commentFlag) {
					try (FileOutputStream outputStream = new FileOutputStream(RESULT_FILE)) {
						resultWorkbook.write(outputStream);
						System.out.println("Diff excel has been generated!");
					}
				}
			} else 
				System.out.println("No diff found!");
		} catch (Exception e) {
			e.printStackTrace(System.out);
		}
	}

	private static List<Future<CallableValue>> processAllSheets(XSSFWorkbook resultWorkbook, XSSFWorkbook workbook2) throws Exception {
		List<SheetProcessorTask> tasks = new ArrayList<>();
		
		Consumer<Sheet> consumer = sheet1 -> {
			XSSFSheet sheet2 = (XSSFSheet) workbook2.getSheet(sheet1.getSheetName());

			if(sheet2 == null) {
				System.out.println(String.format("Sheet[%s] doesn't exist in workbook[%s]", sheet1.getSheetName(), FILE_NAME2));
			} else
				tasks.add(new SheetProcessorTask((XSSFSheet) sheet1, sheet2, commentFlag));
		};

		resultWorkbook.forEach(consumer);
		
		int effectiveFixedThreadPoolSize = Math.min(Runtime.getRuntime().availableProcessors(), tasks.size());
		ExecutorService executor = Executors.newFixedThreadPool(effectiveFixedThreadPoolSize);
		List<Future<CallableValue>> futures = executor.invokeAll(tasks);
		executor.shutdown();
		
		return futures;
	}

}