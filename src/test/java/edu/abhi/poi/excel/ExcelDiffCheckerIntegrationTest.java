package edu.abhi.poi.excel;


import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.util.concurrent.Callable;

import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ExcelDiffCheckerIntegrationTest {

    private XSSFWorkbook workbookOne;
    private XSSFWorkbook workbookTwo;

    // Test setup with existing file paths
    @BeforeEach
    void setUp() throws Exception {
        workbookOne = new XSSFWorkbook(new FileInputStream("C:\\Users\\anujr\\Downloads\\excel1.xlsx"));
        workbookTwo = new XSSFWorkbook(new FileInputStream("C:\\Users\\anujr\\Downloads\\excel2.xlsx"));
    }

    // Test calling two Excel files
    @Test
    void testSheetProcessorTask() throws Exception {
        XSSFSheet sheetOne = workbookOne.getSheetAt(0);
        XSSFSheet sheetTwo = workbookTwo.getSheetAt(0);
        
        // Create a task for processing the sheets
        Callable<CallableValue> sheetTask = new SheetProcessorTask(sheetOne, sheetTwo, false);
        
        // Execute the task
        CallableValue taskResult = sheetTask.call();
        
        // Assert that the result is not null
        assertNotNull(taskResult);

        System.out.println("Test passed successfully!");

        
        // Additional assertions can be made on taskResult to check for expected behavior
        // For example:
        // assertTrue(taskResult.getDiffFlag(), "Differences should be detected");
    }
}
