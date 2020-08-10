package edu.abhi.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.function.Consumer;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
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
	private static boolean remarksOnly = false;
	private static String specifiedSheets = null;
	private static String FILE_NAME1, FILE_NAME2;
	private static CommandLine cmd = null;

	public static void main(String[] args) {
		
		processOptions(args);
		
		String file1path = Paths.get(FILE_NAME1).getFileName().toString(), file2path = Paths.get(FILE_NAME2).getFileName().toString(); 

		String RESULT_FILE = file1path.substring(0, file1path.lastIndexOf(".")) + " vs " + file2path;

		File resultFile = new File(RESULT_FILE);

		Utility.deleteIfExists(resultFile);

		try(XSSFWorkbook resultWorkbook = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME1)));
				XSSFWorkbook workbook2 = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME2)))) {
			
			List<SheetProcessorTask> tasks = specifiedSheets != null? 
					processSpecifiedSheets(resultWorkbook, workbook2) :	processAllSheets(resultWorkbook, workbook2);

			diffFound = CallableValue.analyzeResult(executeAllTasks(tasks));

			if(diffFound) {
				if(!remarksOnly) {
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

	private static List<Future<CallableValue>> executeAllTasks(List<SheetProcessorTask> tasks) throws Exception {
		int effectiveFixedThreadPoolSize = Math.min(Runtime.getRuntime().availableProcessors(), tasks.size());
		ExecutorService executor = Executors.newFixedThreadPool(effectiveFixedThreadPoolSize);
		List<Future<CallableValue>> futures = executor.invokeAll(tasks);
		executor.shutdown();
		
		return futures;
	}

	private static void processOptions(String[] args) {
		Options options = new Options();

        Option base = new Option("b", "base", true, "base file path");
        base.setRequired(true);
        options.addOption(base);

        Option target = new Option("t", "target", true, "target file path");
        target.setRequired(true);
        options.addOption(target);
        
        options.addOption(new Option("s", "sheets", true, "Comma separated sheets"));
        
        options.addOption(new Option("r", "remarks", false, "display only remarks"));

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();

        try {
            cmd = parser.parse(options, args);
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("utility-name", options);

            System.exit(1);
        }
        
        FILE_NAME1 = cmd.getOptionValue("b");
		FILE_NAME2 = cmd.getOptionValue("t");
		specifiedSheets = cmd.getOptionValue("s");
		remarksOnly = cmd.hasOption("r");
	}

	private static List<SheetProcessorTask> processAllSheets(XSSFWorkbook resultWorkbook, XSSFWorkbook workbook2) {
		List<SheetProcessorTask> tasks = new ArrayList<>();
		
		Consumer<Sheet> consumer = sheet1 -> {
			XSSFSheet sheet2 = (XSSFSheet) workbook2.getSheet(sheet1.getSheetName());

			if(sheet2 == null) {
				System.out.println(String.format("Sheet[%s] doesn't exist in workbook[%s]", sheet1.getSheetName(), FILE_NAME2));
			} else
				tasks.add(new SheetProcessorTask((XSSFSheet) sheet1, sheet2, remarksOnly));
		};

		resultWorkbook.forEach(consumer);
		
		return tasks;
	}
	
	private static List<SheetProcessorTask> processSpecifiedSheets(XSSFWorkbook resultWorkbook, XSSFWorkbook workbook2) {
		List<SheetProcessorTask> tasks = new ArrayList<>();
		
		List<String> sheets = Arrays.asList(specifiedSheets.split(","));
		
		Consumer<String> consumer = sheetName -> {
			XSSFSheet sheet1 = (XSSFSheet) resultWorkbook.getSheet(sheetName);
			XSSFSheet sheet2 = (XSSFSheet) workbook2.getSheet(sheetName);

			if(sheet1 == null || sheet2 == null ) {
				System.out.println(String.format("Sheet[%s] doesn't exist in both workbooks", sheetName));
			} else
				tasks.add(new SheetProcessorTask((XSSFSheet) sheet1, sheet2, remarksOnly));
		};

		sheets.forEach(consumer);
		
		return tasks;
	}

}